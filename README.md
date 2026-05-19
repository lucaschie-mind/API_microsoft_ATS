# Meeting Scheduler — protótipo

Backend Python/FastAPI + Postgres (Railway) + Frontend React/Vite.

## Arquivos

```
backend/
  main.py          ← tudo: config, banco, models, endpoints
  requirements.txt
  railway.json
  .env.example

frontend/
  src/
    api.js         ← MSAL + Graph + chamadas ao backend
    App.jsx        ← tela única com agendamento e histórico
    main.jsx
  index.html
  package.json
  vite.config.js
  railway.json
  .env.example
```

---

## Rodar localmente

### Banco (Docker)
```bash
docker run --name ms-db \
  -e POSTGRES_PASSWORD=postgres \
  -e POSTGRES_DB=meeting_scheduler \
  -p 5432:5432 -d postgres:16
```

### Backend
```bash
cd backend
python -m venv venv && source venv/bin/activate
pip install -r requirements.txt
cp .env.example .env   # preencher CLIENT_ID e TENANT_ID do Azure
uvicorn main:app --reload --port 8000
```

### Frontend
```bash
cd frontend
npm install
cp .env.example .env   # preencher VITE_AZURE_CLIENT_ID e VITE_AZURE_TENANT_ID
npm run dev            # abre localhost:5173
```

> O Vite redireciona `/api` → `localhost:8000` automaticamente em dev.

---

## Subir no Railway

### 1. Criar projeto + banco
- railway.app → New Project → Add PostgreSQL
- Copiar a DATABASE_URL gerada

### 2. Backend
- Add service → GitHub → Root Directory: `backend`
- Variables:
  ```
  DATABASE_URL     = (do Railway — trocar postgresql:// por postgresql+asyncpg://)
  ENVIRONMENT      = production
  TIMEZONE         = America/Sao_Paulo
  ALLOWED_ORIGINS  = https://SEU-FRONTEND.up.railway.app
  AZURE_TENANT_ID  = (do Azure Portal)
  AZURE_CLIENT_ID  = (do Azure Portal)
  ```
- Aguardar deploy → testar: `https://SEU-BACKEND.up.railway.app/api/health`

### 3. Adicionar Redirect URI no Azure
- Azure Portal → App registrations → Authentication
- Adicionar: `https://SEU-FRONTEND.up.railway.app`

### 4. Frontend
- Add service → GitHub → Root Directory: `frontend`
- Variables (definir ANTES do deploy):
  ```
  VITE_AZURE_CLIENT_ID = (do Azure Portal)
  VITE_AZURE_TENANT_ID = (do Azure Portal)
  VITE_BACKEND_URL     = https://SEU-BACKEND.up.railway.app
  ```
- Aguardar build → compartilhar URL com testadores

### 5. CORS
- Backend → Variables → atualizar `ALLOWED_ORIGINS` com a URL real do frontend

---

## Banco de dados

As tabelas `meetings` e `participants` são criadas automaticamente no primeiro start.

```sql
-- Ver reuniões:
SELECT subject, start_time, status, teams_link FROM meetings ORDER BY start_time DESC;

-- Ver participantes:
SELECT m.subject, p.email FROM meetings m JOIN participants p ON p.meeting_id = m.id;
```
