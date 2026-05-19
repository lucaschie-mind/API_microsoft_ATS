# main.py
# Meeting Scheduler — backend completo em um arquivo
# FastAPI + SQLAlchemy + Microsoft Graph (modo Application/Pro)

import os
import uuid
import httpx
import pytz
from datetime import datetime, timedelta
from contextlib import asynccontextmanager

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy.ext.asyncio import create_async_engine, async_sessionmaker, AsyncSession
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy import String, DateTime, Boolean, Integer, Text, ForeignKey, select
from azure.identity.aio import ClientSecretCredential

# ─── Config ──────────────────────────────────────────────────────────────────

def _fix_db_url(url: str) -> str:
    """Railway injeta postgres:// — SQLAlchemy async precisa de postgresql+asyncpg://"""
    if url.startswith("postgres://"):
        return url.replace("postgres://", "postgresql+asyncpg://", 1)
    if url.startswith("postgresql://") and "+asyncpg" not in url:
        return url.replace("postgresql://", "postgresql+asyncpg://", 1)
    return url

DATABASE_URL   = _fix_db_url(os.environ.get("DATABASE_URL", "postgresql+asyncpg://postgres:postgres@localhost:5432/meeting_scheduler"))
ALLOWED_ORIGINS = os.environ.get("ALLOWED_ORIGINS", "http://localhost:5173,http://localhost:3000").split(",")
TIMEZONE        = os.environ.get("TIMEZONE", "America/Sao_Paulo")
ENVIRONMENT     = os.environ.get("ENVIRONMENT", "development")

# Azure — necessário apenas no modo Pro (Application)
AZURE_TENANT_ID     = os.environ.get("AZURE_TENANT_ID", "")
AZURE_CLIENT_ID     = os.environ.get("AZURE_CLIENT_ID", "")
AZURE_CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "")

# Tenant fixo para o protótipo — gerado uma vez e salvo no .env
# Em produção: cada empresa cliente teria seu próprio registro no banco
TENANT_ID = os.environ.get("TENANT_ID", str(uuid.uuid4()))

GRAPH_BASE  = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = "https://graph.microsoft.com/.default"

# ─── Banco de dados ───────────────────────────────────────────────────────────

engine       = create_async_engine(DATABASE_URL, echo=ENVIRONMENT == "development")
SessionLocal = async_sessionmaker(engine, class_=AsyncSession, expire_on_commit=False)

class Base(DeclarativeBase):
    pass

class Meeting(Base):
    __tablename__ = "meetings"

    id:              Mapped[uuid.UUID]   = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    subject:         Mapped[str]         = mapped_column(String(500))
    start_time:      Mapped[datetime]    = mapped_column(DateTime)
    end_time:        Mapped[datetime]    = mapped_column(DateTime)
    duration_minutes:Mapped[int]         = mapped_column(Integer, default=60)
    is_online:       Mapped[bool]        = mapped_column(Boolean, default=True)
    teams_link:      Mapped[str | None]  = mapped_column(Text, nullable=True)
    outlook_link:    Mapped[str | None]  = mapped_column(Text, nullable=True)
    ms_event_id:     Mapped[str | None]  = mapped_column(String(500), nullable=True)
    organizer_email: Mapped[str]         = mapped_column(String(255))
    organizer_name:  Mapped[str]         = mapped_column(String(255))
    status:          Mapped[str]         = mapped_column(String(50), default="scheduled")
    created_at:      Mapped[datetime]    = mapped_column(DateTime, default=datetime.utcnow)

    participants: Mapped[list["Participant"]] = relationship("Participant", back_populates="meeting", cascade="all, delete-orphan")


class Participant(Base):
    __tablename__ = "participants"

    id:           Mapped[uuid.UUID]  = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    meeting_id:   Mapped[uuid.UUID]  = mapped_column(UUID(as_uuid=True), ForeignKey("meetings.id"))
    email:        Mapped[str]        = mapped_column(String(255))
    is_organizer: Mapped[bool]       = mapped_column(Boolean, default=False)

    meeting: Mapped["Meeting"] = relationship("Meeting", back_populates="participants")


# ─── App ─────────────────────────────────────────────────────────────────────

@asynccontextmanager
async def lifespan(app: FastAPI):
    async with engine.begin() as conn:
        await conn.run_sync(Base.metadata.create_all)
    yield
    await engine.dispose()

app = FastAPI(
    title="Meeting Scheduler API",
    version="1.0.0",
    lifespan=lifespan,
    docs_url="/docs" if ENVIRONMENT == "development" else None,
    redoc_url=None,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─── Helpers de slot ─────────────────────────────────────────────────────────

def find_slots(schedules: list[dict], window: dict, duration_min: int, max_slots: int = 3) -> list[dict]:
    """Encontra horários onde todos os participantes estão livres."""
    tz   = pytz.timezone(TIMEZONE)
    now  = datetime.now(tz)
    slots = []

    for day_offset in range(window.get("days_ahead", 14)):
        day = now + timedelta(days=day_offset)
        if day.weekday() >= 5:  # pula fim de semana
            continue

        date_str = day.strftime("%Y-%m-%d")

        for w_start, w_end in [
            (window.get("morning_start", "09:00"),   window.get("morning_end",   "12:00")),
            (window.get("afternoon_start", "13:00"), window.get("afternoon_end", "18:00")),
        ]:
            cursor  = tz.localize(datetime.strptime(f"{date_str} {w_start}", "%Y-%m-%d %H:%M"))
            win_end = tz.localize(datetime.strptime(f"{date_str} {w_end}",   "%Y-%m-%d %H:%M"))

            while cursor + timedelta(minutes=duration_min) <= win_end:
                if cursor > now and _all_free(schedules, cursor, duration_min):
                    slot_end = cursor + timedelta(minutes=duration_min)
                    slots.append({
                        "start": cursor.isoformat(),
                        "end":   slot_end.isoformat(),
                    })
                    if len(slots) >= max_slots:
                        return slots
                cursor += timedelta(minutes=30)

    return slots


def _all_free(schedules: list[dict], start: datetime, duration_min: int) -> bool:
    end = start + timedelta(minutes=duration_min)
    for schedule in schedules:
        for item in schedule.get("scheduleItems", []):
            if item.get("status") not in ("busy", "oof"):
                continue
            raw_s = item.get("start", {}).get("dateTime", "")
            raw_e = item.get("end",   {}).get("dateTime", "")
            if not raw_s or not raw_e:
                continue
            bs = _parse_dt(raw_s)
            be = _parse_dt(raw_e)
            if bs < end and be > start:
                return False
    return True


def _parse_dt(raw: str) -> datetime:
    dt = datetime.fromisoformat(raw.replace("Z", "+00:00"))
    return dt if dt.tzinfo else pytz.utc.localize(dt)


# ─── Graph API (modo Pro) ─────────────────────────────────────────────────────

async def graph_post(path: str, body: dict) -> dict:
    credential = ClientSecretCredential(AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET)
    try:
        token = await credential.get_token(GRAPH_SCOPE)
        async with httpx.AsyncClient() as client:
            r = await client.post(
                f"{GRAPH_BASE}{path}",
                headers={"Authorization": f"Bearer {token.token}", "Content-Type": "application/json"},
                json=body, timeout=15,
            )
            r.raise_for_status()
            return r.json()
    finally:
        await credential.close()


# ─── Endpoints ────────────────────────────────────────────────────────────────

@app.get("/api/health")
async def health():
    return {"status": "ok", "environment": ENVIRONMENT}


# Disponibilidade — usado pelo modo Pro
# No modo Basic, o frontend chama o Graph diretamente via MSAL
@app.post("/api/schedule")
async def get_schedule(payload: dict):
    """Busca free/busy via Graph (modo Pro) e retorna slots disponíveis."""
    emails   = payload.get("emails", [])
    duration = payload.get("duration_minutes", 60)
    window   = payload.get("window", {})

    if not emails:
        raise HTTPException(status_code=400, detail="emails é obrigatório")
    if not all([AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET]):
        raise HTTPException(status_code=501, detail="Credenciais Azure não configuradas — use modo Basic")

    tz  = pytz.timezone(TIMEZONE)
    now = datetime.now(tz)
    end = now + timedelta(days=window.get("days_ahead", 14))

    try:
        result = await graph_post(
            f"/users/{emails[0]}/calendar/getSchedule",
            {
                "schedules": emails,
                "startTime": {"dateTime": now.isoformat(), "timeZone": TIMEZONE},
                "endTime":   {"dateTime": end.isoformat(), "timeZone": TIMEZONE},
                "availabilityViewInterval": 30,
            },
        )
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Erro ao consultar Graph API: {e}")

    return {"slots": find_slots(result.get("value", []), window, duration)}


# Salvar reunião — usado por ambos os modos após o evento ser criado
@app.post("/api/meetings", status_code=201)
async def create_meeting(payload: dict):
    """Salva a reunião no Postgres. O evento já foi criado no Graph pelo frontend (modo Basic)
    ou é criado aqui via Graph antes de salvar (modo Pro)."""

    required = ["subject", "start_time", "end_time", "organizer_email",
                "organizer_name", "attendee_emails"]
    for field in required:
        if not payload.get(field):
            raise HTTPException(status_code=400, detail=f"Campo obrigatório ausente: {field}")

    async with SessionLocal() as db:
        meeting = Meeting(
            subject          = payload["subject"],
            start_time       = datetime.fromisoformat(payload["start_time"]),
            end_time         = datetime.fromisoformat(payload["end_time"]),
            duration_minutes = payload.get("duration_minutes", 60),
            is_online        = payload.get("is_online", True),
            teams_link       = payload.get("teams_link"),
            outlook_link     = payload.get("outlook_link"),
            ms_event_id      = payload.get("ms_event_id"),
            organizer_email  = payload["organizer_email"],
            organizer_name   = payload["organizer_name"],
        )
        db.add(meeting)
        await db.flush()

        for email in dict.fromkeys(payload["attendee_emails"]):
            db.add(Participant(
                meeting_id   = meeting.id,
                email        = email,
                is_organizer = (email == payload["organizer_email"]),
            ))

        await db.commit()
        return _meeting_to_dict(meeting, payload["attendee_emails"])


# Listar reuniões
@app.get("/api/meetings")
async def list_meetings(
    organizer_email: str | None = None,
    page: int = 1,
    limit: int = 20,
):
    async with SessionLocal() as db:
        q = select(Meeting).order_by(Meeting.start_time.desc())
        if organizer_email:
            q = q.where(Meeting.organizer_email == organizer_email)
        meetings = (await db.execute(q.offset((page - 1) * limit).limit(limit))).scalars().all()

        result = []
        for m in meetings:
            parts = (await db.execute(
                select(Participant).where(Participant.meeting_id == m.id)
            )).scalars().all()
            result.append(_meeting_to_dict(m, [p.email for p in parts]))

        return {"meetings": result, "page": page}


# Cancelar reunião
@app.patch("/api/meetings/{meeting_id}/cancel")
async def cancel_meeting(meeting_id: str):
    async with SessionLocal() as db:
        result = await db.execute(select(Meeting).where(Meeting.id == meeting_id))
        meeting = result.scalar_one_or_none()
        if not meeting:
            raise HTTPException(status_code=404, detail="Reunião não encontrada")
        meeting.status = "cancelled"
        await db.commit()
        return {"id": str(meeting.id), "status": "cancelled"}


# ─── Helper de serialização ───────────────────────────────────────────────────

def _meeting_to_dict(meeting: Meeting, attendee_emails: list[str]) -> dict:
    return {
        "id":               str(meeting.id),
        "subject":          meeting.subject,
        "start_time":       meeting.start_time.isoformat(),
        "end_time":         meeting.end_time.isoformat(),
        "duration_minutes": meeting.duration_minutes,
        "is_online":        meeting.is_online,
        "teams_link":       meeting.teams_link,
        "outlook_link":     meeting.outlook_link,
        "ms_event_id":      meeting.ms_event_id,
        "organizer_email":  meeting.organizer_email,
        "organizer_name":   meeting.organizer_name,
        "status":           meeting.status,
        "attendee_emails":  attendee_emails,
        "created_at":       meeting.created_at.isoformat(),
    }
