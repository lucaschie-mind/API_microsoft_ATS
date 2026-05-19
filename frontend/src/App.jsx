// src/App.jsx
import { useState, useEffect, useCallback } from 'react'
import {
  getMe, getAccount, logout,
  getFreeBusy, findSlots, createGraphEvent,
  saveMeeting, listMeetings, cancelMeeting,
} from './api.js'

// ─── Formatadores ─────────────────────────────────────────────────────────────

function fmtLong(dt) {
  return new Date(dt).toLocaleString('pt-BR', {
    weekday: 'long', day: '2-digit', month: 'long',
    hour: '2-digit', minute: '2-digit',
  })
}

function fmtShort(dt) {
  return new Date(dt).toLocaleString('pt-BR', {
    weekday: 'short', day: '2-digit', month: 'short',
    hour: '2-digit', minute: '2-digit',
  })
}

function fmtTime(dt) {
  return new Date(dt).toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })
}

// ─── Estilos ─────────────────────────────────────────────────────────────────

const s = {
  page:      { maxWidth: 620, margin: '0 auto', padding: '2rem 1rem', fontFamily: 'system-ui, sans-serif', color: '#1a1a1a' },
  card:      { background: '#fff', border: '1px solid #e5e7eb', borderRadius: 12, padding: '1.25rem', marginBottom: 12 },
  cardTitle: { fontSize: 11, fontWeight: 700, color: '#9ca3af', textTransform: 'uppercase', letterSpacing: '.06em', marginBottom: 12 },
  label:     { display: 'block', fontSize: 12, color: '#6b7280', marginBottom: 4 },
  field:     { marginBottom: 10 },
  row2:      { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 },
  row3:      { display: 'grid', gridTemplateColumns: '80px 1fr 1fr', gap: 12 },
  input:     { width: '100%', padding: '8px 10px', border: '1px solid #d1d5db', borderRadius: 8, fontSize: 13, boxSizing: 'border-box' },
  btnBlue:   { width: '100%', padding: 12, fontSize: 14, fontWeight: 600, background: '#2563eb', color: '#fff', border: 'none', borderRadius: 10, cursor: 'pointer' },
  btnGreen:  { width: '100%', padding: 12, fontSize: 14, fontWeight: 600, background: '#059669', color: '#fff', border: 'none', borderRadius: 10, cursor: 'pointer', marginTop: 10 },
  btnGhost:  { background: 'none', border: '1px solid #e5e7eb', borderRadius: 6, padding: '4px 10px', fontSize: 12, cursor: 'pointer', color: '#6b7280' },
  tag:       { display: 'inline-flex', alignItems: 'center', gap: 4, background: '#eff6ff', color: '#1d4ed8', borderRadius: 6, padding: '3px 8px', fontSize: 12 },
  tagX:      { background: 'none', border: 'none', cursor: 'pointer', color: 'inherit', fontSize: 14, padding: 0, lineHeight: 1 },
  error:     { fontSize: 13, color: '#dc2626', margin: '8px 0' },
  timeRow:   { display: 'flex', gap: 4, alignItems: 'center' },
  sep:       { fontSize: 11, color: '#9ca3af' },
}

// ─── App ─────────────────────────────────────────────────────────────────────

export default function App() {
  // Auth
  const [user,    setUser]    = useState(null)
  const [loading, setLoading] = useState(false)
  const [error,   setError]   = useState(null)

  // Form
  const [subject,        setSubject]        = useState('')
  const [duration,       setDuration]       = useState(60)
  const [guestInput,     setGuestInput]     = useState('')
  const [guests,         setGuests]         = useState([])
  const [isOnline,       setIsOnline]       = useState(true)
  const [daysAhead,      setDaysAhead]      = useState(14)
  const [morningStart,   setMorningStart]   = useState('09:00')
  const [morningEnd,     setMorningEnd]     = useState('12:00')
  const [afternoonStart, setAfternoonStart] = useState('13:00')
  const [afternoonEnd,   setAfternoonEnd]   = useState('18:00')

  // Resultados
  const [slots,        setSlots]        = useState([])
  const [selectedSlot, setSelectedSlot] = useState(null)
  const [meeting,      setMeeting]      = useState(null)

  // Histórico
  const [showHistory,  setShowHistory]  = useState(false)
  const [history,      setHistory]      = useState([])
  const [histLoading,  setHistLoading]  = useState(false)

  // Inicializar — verificar se já tem conta logada
  useEffect(() => {
    const account = getAccount()
    if (account) {
      getMe()
        .then(me => setUser({ name: me.displayName, email: me.mail || me.userPrincipalName }))
        .catch(() => {})
    }
  }, [])

  const handleLogout = async () => {
    await logout()
    setUser(null)
    setSlots([])
    setMeeting(null)
  }

  // ── Buscar slots ─────────────────────────────────────────────────────────

  const handleSearch = useCallback(async () => {
    if (!subject.trim()) { setError('Informe o título da reunião'); return }
    if (!guests.length)  { setError('Adicione ao menos um convidado'); return }

    setLoading(true)
    setError(null)
    setSlots([])
    setSelectedSlot(null)

    try {
      let me = user
      if (!me) {
        const data = await getMe()
        me = { name: data.displayName, email: data.mail || data.userPrincipalName }
        setUser(me)
      }

      const allEmails = [me.email, ...guests]
      const schedules = await getFreeBusy(allEmails, daysAhead)
      const found     = findSlots(schedules, { daysAhead, morningStart, morningEnd, afternoonStart, afternoonEnd }, duration)

      if (!found.length) setError('Nenhum horário disponível encontrado. Tente ampliar a janela.')
      else setSlots(found)
    } catch (e) {
      setError(e.message)
    } finally {
      setLoading(false)
    }
  }, [user, subject, guests, duration, daysAhead, morningStart, morningEnd, afternoonStart, afternoonEnd])

  // ── Confirmar ────────────────────────────────────────────────────────────

  const handleConfirm = useCallback(async () => {
    if (!selectedSlot || !user) return
    setLoading(true)
    setError(null)

    try {
      const allEmails = [user.email, ...guests]

      // 1. Criar no Graph via MSAL
      const event = await createGraphEvent(
        subject, selectedSlot.start, selectedSlot.end, allEmails, isOnline
      )

      // 2. Salvar no Postgres
      const saved = await saveMeeting({
        subject,
        start_time:       selectedSlot.start.toISOString(),
        end_time:         selectedSlot.end.toISOString(),
        duration_minutes: duration,
        is_online:        isOnline,
        teams_link:       event.onlineMeeting?.joinUrl,
        outlook_link:     event.webLink,
        ms_event_id:      event.id,
        organizer_email:  user.email,
        organizer_name:   user.name,
        attendee_emails:  allEmails,
      })

      setMeeting(saved)
      setSlots([])
    } catch (e) {
      setError(e.message)
    } finally {
      setLoading(false)
    }
  }, [selectedSlot, user, subject, guests, isOnline, duration])

  // ── Histórico ────────────────────────────────────────────────────────────

  const loadHistory = useCallback(async () => {
    if (!user) return
    setHistLoading(true)
    try {
      const data = await listMeetings(user.email)
      setHistory(data.meetings || [])
    } catch (e) {
      setError(e.message)
    } finally {
      setHistLoading(false)
    }
  }, [user])

  const toggleHistory = () => {
    if (!showHistory) loadHistory()
    setShowHistory(v => !v)
  }

  const handleCancel = async (id) => {
    if (!confirm('Cancelar esta reunião?')) return
    try {
      await cancelMeeting(id)
      loadHistory()
    } catch (e) {
      setError(e.message)
    }
  }

  // ── Render ────────────────────────────────────────────────────────────────

  return (
    <div style={{ background: '#f4f6f9', minHeight: '100vh', padding: '2rem 1rem' }}>
      <div style={s.page}>

        {/* Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 24 }}>
          <div>
            <h1 style={{ fontSize: 22, fontWeight: 700, color: '#0f2d5e', margin: 0 }}>Meeting Scheduler</h1>
            <p style={{ fontSize: 13, color: '#6b7280', margin: '2px 0 0' }}>Agendamento inteligente com Microsoft 365</p>
          </div>
          {user && (
            <div style={{ textAlign: 'right' }}>
              <p style={{ fontSize: 12, color: '#6b7280', margin: 0 }}>{user.name}</p>
              <button onClick={handleLogout} style={s.btnGhost}>Sair</button>
            </div>
          )}
        </div>

        {/* Reunião criada — estado de sucesso */}
        {meeting ? (
          <div style={{ ...s.card, borderColor: '#10b981', borderWidth: 2 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, color: '#059669', fontWeight: 700, fontSize: 16, marginBottom: 16 }}>
              <span style={{ fontSize: 20 }}>✓</span> Reunião criada com sucesso
            </div>
            {[
              ['Título',        meeting.subject],
              ['Data',          `${fmtLong(meeting.start_time)} – ${fmtTime(meeting.end_time)}`],
              ['Participantes', meeting.attendee_emails?.join(', ')],
            ].map(([label, value]) => (
              <div key={label} style={{ display: 'flex', gap: 12, marginBottom: 8, fontSize: 13 }}>
                <span style={{ minWidth: 90, fontWeight: 600, color: '#6b7280', flexShrink: 0 }}>{label}</span>
                <span>{value}</span>
              </div>
            ))}
            {meeting.teams_link && (
              <div style={{ display: 'flex', gap: 12, marginBottom: 8, fontSize: 13 }}>
                <span style={{ minWidth: 90, fontWeight: 600, color: '#6b7280', flexShrink: 0 }}>Teams</span>
                <a href={meeting.teams_link} target="_blank" rel="noreferrer"
                   style={{ color: '#2563eb', wordBreak: 'break-all' }}>{meeting.teams_link}</a>
              </div>
            )}
            {meeting.outlook_link && (
              <a href={meeting.outlook_link} target="_blank" rel="noreferrer"
                 style={{ fontSize: 13, color: '#2563eb', display: 'inline-block', marginTop: 4 }}>
                Ver no Outlook ↗
              </a>
            )}
            <button onClick={() => { setMeeting(null); setSubject(''); setGuests([]) }}
                    style={{ ...s.btnGhost, marginTop: 16, display: 'block' }}>
              Agendar outra reunião
            </button>
          </div>
        ) : (
          <>
            {/* Detalhes */}
            <div style={s.card}>
              <p style={s.cardTitle}>Detalhes</p>
              <div style={s.field}>
                <label style={s.label}>Título da reunião *</label>
                <input style={s.input} value={subject} onChange={e => setSubject(e.target.value)}
                       placeholder="Ex: Alinhamento Q3" />
              </div>
              <div style={s.row2}>
                <div style={s.field}>
                  <label style={s.label}>Duração</label>
                  <select style={s.input} value={duration} onChange={e => setDuration(Number(e.target.value))}>
                    {[30, 45, 60, 90, 120].map(d => <option key={d} value={d}>{d} min</option>)}
                  </select>
                </div>
                <div style={{ ...s.field, display: 'flex', alignItems: 'center', gap: 8, paddingTop: 18 }}>
                  <input type="checkbox" id="online" checked={isOnline}
                         onChange={e => setIsOnline(e.target.checked)} style={{ width: 15, height: 15 }} />
                  <label htmlFor="online" style={{ fontSize: 13, cursor: 'pointer' }}>Criar link Teams</label>
                </div>
              </div>
            </div>

            {/* Convidados */}
            <div style={s.card}>
              <p style={s.cardTitle}>Convidados</p>
              <div style={{ display: 'flex', gap: 8, marginBottom: 8 }}>
                <input style={{ ...s.input, flex: 1 }} value={guestInput}
                       onChange={e => { setGuestInput(e.target.value); setError(null) }}
                       onKeyDown={e => {
                         if (e.key !== 'Enter') return
                         const email = guestInput.trim().toLowerCase()
                         if (email && !guests.includes(email)) setGuests(g => [...g, email])
                         setGuestInput('')
                       }}
                       placeholder="email@empresa.com — Enter para adicionar" />
                <button onClick={() => {
                  const email = guestInput.trim().toLowerCase()
                  if (email && !guests.includes(email)) setGuests(g => [...g, email])
                  setGuestInput('')
                }} style={{ ...s.btnGhost, padding: '8px 14px', fontSize: 18 }}>+</button>
              </div>
              {guests.length > 0 && (
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                  {guests.map(email => (
                    <span key={email} style={s.tag}>
                      {email}
                      <button style={s.tagX} onClick={() => setGuests(g => g.filter(x => x !== email))}>×</button>
                    </span>
                  ))}
                </div>
              )}
            </div>

            {/* Janela */}
            <div style={s.card}>
              <p style={s.cardTitle}>Janela de horários</p>
              <div style={s.row3}>
                <div style={s.field}>
                  <label style={s.label}>Dias</label>
                  <input type="number" style={s.input} value={daysAhead} min={1} max={60}
                         onChange={e => setDaysAhead(Number(e.target.value))} />
                </div>
                <div style={s.field}>
                  <label style={s.label}>Manhã</label>
                  <div style={s.timeRow}>
                    <input type="time" style={s.input} value={morningStart} onChange={e => setMorningStart(e.target.value)} />
                    <span style={s.sep}>–</span>
                    <input type="time" style={s.input} value={morningEnd} onChange={e => setMorningEnd(e.target.value)} />
                  </div>
                </div>
                <div style={s.field}>
                  <label style={s.label}>Tarde</label>
                  <div style={s.timeRow}>
                    <input type="time" style={s.input} value={afternoonStart} onChange={e => setAfternoonStart(e.target.value)} />
                    <span style={s.sep}>–</span>
                    <input type="time" style={s.input} value={afternoonEnd} onChange={e => setAfternoonEnd(e.target.value)} />
                  </div>
                </div>
              </div>
            </div>

            {error && <p style={s.error}>{error}</p>}

            <button onClick={handleSearch} disabled={loading} style={{ ...s.btnBlue, opacity: loading ? .6 : 1 }}>
              {loading && !slots.length ? 'Buscando horários...' : 'Verificar disponibilidade'}
            </button>

            {/* Slots */}
            {slots.length > 0 && (
              <div style={{ ...s.card, marginTop: 12 }}>
                <p style={s.cardTitle}>Horários disponíveis</p>
                {slots.map((slot, i) => {
                  const active = selectedSlot === slot
                  return (
                    <button key={i} onClick={() => setSelectedSlot(slot)} style={{
                      display: 'flex', justifyContent: 'space-between', alignItems: 'center',
                      width: '100%', padding: '12px 14px', marginBottom: 6,
                      background: active ? '#eff6ff' : '#f9fafb',
                      border: `${active ? 2 : 1}px solid ${active ? '#2563eb' : '#e5e7eb'}`,
                      borderRadius: 8, cursor: 'pointer', textAlign: 'left',
                    }}>
                      <div>
                        <div style={{ fontSize: 14, fontWeight: 600 }}>{fmtLong(slot.start)}</div>
                        <div style={{ fontSize: 12, color: '#9ca3af', marginTop: 2 }}>
                          até {fmtTime(slot.end)} · {duration} min
                        </div>
                      </div>
                      {active && <span style={{ color: '#2563eb', fontSize: 18 }}>✓</span>}
                    </button>
                  )
                })}
                {selectedSlot && (
                  <button onClick={handleConfirm} disabled={loading}
                          style={{ ...s.btnGreen, opacity: loading ? .6 : 1 }}>
                    {loading ? 'Criando reunião...' : 'Confirmar e criar reunião'}
                  </button>
                )}
              </div>
            )}
          </>
        )}

        {/* Histórico — colapsável */}
        {user && (
          <div style={{ marginTop: 24 }}>
            <button onClick={toggleHistory} style={{
              ...s.btnGhost, width: '100%', padding: '10px',
              display: 'flex', justifyContent: 'space-between', alignItems: 'center',
            }}>
              <span style={{ fontWeight: 600, fontSize: 13 }}>Histórico de reuniões</span>
              <span style={{ fontSize: 16 }}>{showHistory ? '▲' : '▼'}</span>
            </button>

            {showHistory && (
              <div style={{ marginTop: 8 }}>
                {histLoading ? (
                  <p style={{ fontSize: 13, color: '#9ca3af', padding: '1rem', textAlign: 'center' }}>Carregando...</p>
                ) : !history.length ? (
                  <p style={{ fontSize: 13, color: '#9ca3af', padding: '1rem', textAlign: 'center' }}>Nenhuma reunião ainda.</p>
                ) : history.map(m => {
                  const statusColor = m.status === 'cancelled' ? '#dc2626' : m.status === 'completed' ? '#7c3aed' : '#059669'
                  const statusLabel = { scheduled: 'Agendada', cancelled: 'Cancelada', completed: 'Concluída' }[m.status] || m.status
                  return (
                    <div key={m.id} style={{ ...s.card, marginBottom: 8 }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 6 }}>
                        <span style={{ fontWeight: 600, fontSize: 14 }}>{m.subject}</span>
                        <span style={{ fontSize: 11, fontWeight: 600, color: statusColor, background: `${statusColor}18`,
                                       borderRadius: 6, padding: '2px 8px', whiteSpace: 'nowrap' }}>
                          {statusLabel}
                        </span>
                      </div>
                      <p style={{ fontSize: 12, color: '#6b7280', margin: '0 0 6px' }}>
                        {fmtShort(m.start_time)} – {fmtTime(m.end_time)} · {m.duration_minutes} min
                      </p>
                      <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
                        {m.teams_link && (
                          <a href={m.teams_link} target="_blank" rel="noreferrer"
                             style={{ fontSize: 12, color: '#2563eb' }}>Abrir Teams ↗</a>
                        )}
                        {m.status === 'scheduled' && (
                          <button onClick={() => handleCancel(m.id)}
                                  style={{ fontSize: 12, padding: '3px 10px', background: 'none',
                                           border: '1px solid #fca5a5', borderRadius: 6, color: '#dc2626', cursor: 'pointer' }}>
                            Cancelar
                          </button>
                        )}
                      </div>
                    </div>
                  )
                })}
                <button onClick={loadHistory} style={{ ...s.btnGhost, width: '100%', marginTop: 4 }}>↻ Atualizar</button>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  )
}
