// src/api.js
// Todas as chamadas ao Microsoft Graph (via MSAL) e ao backend

import { PublicClientApplication, InteractionRequiredAuthError } from '@azure/msal-browser'

const CLIENT_ID  = import.meta.env.VITE_AZURE_CLIENT_ID
const TENANT_ID  = import.meta.env.VITE_AZURE_TENANT_ID
const BACKEND    = import.meta.env.VITE_BACKEND_URL ? `${import.meta.env.VITE_BACKEND_URL}/api` : '/api'
const TZ         = 'America/Sao_Paulo'

// ─── Helpers de fuso ──────────────────────────────────────────────────────────

/**
 * Formata um Date como "YYYY-MM-DDTHH:MM:SS" no fuso de Brasília,
 * SEM converter para UTC. É o formato que o Graph API espera junto com timeZone.
 */
function toBrasiliaISO(date) {
  return new Date(date).toLocaleString('sv-SE', { timeZone: TZ }).replace(' ', 'T')
}

/**
 * Retorna a data local de Brasília como "YYYY-MM-DD"
 * (evita virar dia anterior ao converter para UTC)
 */
function toBrasiliaDateStr(date) {
  return toBrasiliaISO(date).slice(0, 10)
}

/**
 * "Agora" no horário de Brasília como objeto Date corrigido.
 * Usado para comparações no slot finder.
 */
function nowBrasilia() {
  const now = new Date()
  // Offset em minutos de Brasília em relação ao UTC (-180 = UTC-3)
  const offset = -180
  const utc = now.getTime() + now.getTimezoneOffset() * 60000
  return new Date(utc + offset * 60000)
}

// ─── MSAL ─────────────────────────────────────────────────────────────────────

const msal = new PublicClientApplication({
  auth: {
    clientId:    CLIENT_ID,
    authority:   `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: window.location.origin,
  },
  cache: { cacheLocation: 'sessionStorage' },
})

const SCOPES = ['Calendars.ReadWrite', 'User.Read', 'User.ReadBasic.All', 'OnlineMeetings.ReadWrite']

export async function getToken() {
  await msal.initialize()
  const accounts = msal.getAllAccounts()
  if (!accounts.length) {
    const r = await msal.loginPopup({ scopes: SCOPES })
    return r.accessToken
  }
  try {
    const r = await msal.acquireTokenSilent({ scopes: SCOPES, account: accounts[0] })
    return r.accessToken
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError) {
      const r = await msal.acquireTokenPopup({ scopes: SCOPES, account: accounts[0] })
      return r.accessToken
    }
    throw e
  }
}

export async function logout() {
  await msal.initialize()
  const account = msal.getAllAccounts()[0]
  if (account) await msal.logoutPopup({ account })
}

export function getAccount() {
  return msal.getAllAccounts()[0] || null
}

// ─── Graph ────────────────────────────────────────────────────────────────────

async function graph(endpoint, options = {}) {
  const token = await getToken()
  const res = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(options.headers || {}),
    },
  })
  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(err.error?.message || `Graph ${res.status}`)
  }
  return res.status === 204 ? null : res.json()
}

export async function getMe() {
  return graph('/me?$select=displayName,mail,userPrincipalName')
}

export async function getFreeBusy(emails, daysAhead = 14) {
  const now = new Date()
  const end = new Date(now)
  end.setDate(now.getDate() + daysAhead)

  const result = await graph('/me/calendar/getSchedule', {
    method: 'POST',
    body: JSON.stringify({
      schedules: emails,
      // Usar toBrasiliaISO garante que o datetime enviado é Brasília, não UTC
      startTime: { dateTime: toBrasiliaISO(now), timeZone: TZ },
      endTime:   { dateTime: toBrasiliaISO(end), timeZone: TZ },
      availabilityViewInterval: 30,
    }),
  })
  return result.value
}

export async function createGraphEvent(subject, start, end, attendeeEmails, isOnline) {
  return graph('/me/events', {
    method: 'POST',
    body: JSON.stringify({
      subject,
      // toBrasiliaISO garante que o evento é criado no horário certo
      start: { dateTime: toBrasiliaISO(start), timeZone: TZ },
      end:   { dateTime: toBrasiliaISO(end),   timeZone: TZ },
      attendees: attendeeEmails.map(email => ({
        emailAddress: { address: email }, type: 'required',
      })),
      isOnlineMeeting: isOnline,
      onlineMeetingProvider: isOnline ? 'teamsForBusiness' : undefined,
    }),
  })
}

// ─── Slot finder ─────────────────────────────────────────────────────────────

export function findSlots(schedules, window, durationMin, max = 3) {
  const { daysAhead = 14, morningStart = '09:00', morningEnd = '12:00',
          afternoonStart = '13:00', afternoonEnd = '18:00' } = window
  const slots = []
  const now   = nowBrasilia() // "agora" no horário de Brasília

  for (let d = 0; d < daysAhead; d++) {
    const date = new Date(now)
    date.setDate(now.getDate() + d)
    if ([0, 6].includes(date.getDay())) continue

    // Data no fuso de Brasília — evita virar dia anterior em UTC
    const ds = toBrasiliaDateStr(date)

    for (const [ws, we] of [[morningStart, morningEnd], [afternoonStart, afternoonEnd]]) {
      // Criar os cursores já no horário de Brasília
      let cursor  = new Date(`${ds}T${ws}:00`)
      const winEnd = new Date(`${ds}T${we}:00`)

      while (cursor.getTime() + durationMin * 60000 <= winEnd.getTime()) {
        if (cursor > now) {
          const slotEnd = new Date(cursor.getTime() + durationMin * 60000)
          const free = schedules.every(s =>
            !s.scheduleItems?.some(item => {
              if (!['busy', 'oof'].includes(item.status)) return false
              // O Graph retorna os horários dos items em UTC — converter para comparar
              const bs = new Date(item.start.dateTime + (item.start.dateTime.endsWith('Z') ? '' : 'Z'))
              const be = new Date(item.end.dateTime   + (item.end.dateTime.endsWith('Z')   ? '' : 'Z'))
              // cursor e slotEnd estão no horário local do browser — converter para UTC para comparar
              const cursorUTC  = new Date(cursor.getTime()  + (now.getTimezoneOffset() * 60000 * -1) - (180 * 60000 * -1))
              const slotEndUTC = new Date(slotEnd.getTime() + (now.getTimezoneOffset() * 60000 * -1) - (180 * 60000 * -1))
              return bs < slotEndUTC && be > cursorUTC
            })
          )
          if (free) {
            slots.push({ start: new Date(cursor), end: slotEnd })
            if (slots.length >= max) return slots
          }
        }
        cursor = new Date(cursor.getTime() + 30 * 60000)
      }
    }
  }
  return slots
}

// ─── Backend (Postgres) ───────────────────────────────────────────────────────

export async function saveMeeting(data) {
  const res = await fetch(`${BACKEND}/meetings`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  })
  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(err.detail || `Erro ao salvar: ${res.status}`)
  }
  return res.json()
}

export async function listMeetings(organizerEmail, page = 1) {
  const params = new URLSearchParams({ page, limit: 20 })
  if (organizerEmail) params.set('organizer_email', organizerEmail)
  const res = await fetch(`${BACKEND}/meetings?${params}`)
  if (!res.ok) throw new Error(`Erro ao carregar histórico: ${res.status}`)
  return res.json()
}

export async function cancelMeeting(id) {
  const res = await fetch(`${BACKEND}/meetings/${id}/cancel`, { method: 'PATCH' })
  if (!res.ok) throw new Error(`Erro ao cancelar: ${res.status}`)
  return res.json()
}
