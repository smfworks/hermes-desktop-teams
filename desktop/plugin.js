/**
 * Hermes Desktop — Microsoft Teams inbox (SMF Works).
 * Folder name must equal id (`hermes-teams-inbox`).
 * Backend: ~/.hermes/plugins/hermes-teams-inbox (Azure CLI Graph token).
 */
import {
  Badge,
  Button,
  Codicon,
  EmptyState,
  ErrorState,
  GlyphSpinner,
  PALETTE_AREA,
  ROUTES_AREA,
  SIDEBAR_NAV_AREA,
  ScrollArea,
  Separator,
  Tip,
  cn,
  haptic,
  host,
  usePluginI18n,
  useQuery
} from '@hermes/plugin-sdk'
import { useState } from 'react'
import { jsx, jsxs } from 'react/jsx-runtime'

const ID = 'hermes-teams-inbox'
const ROUTE = '/teams'
const TEAMS_WEB_HOSTS = [
  'teams.microsoft.com',
  'teams.microsoft.us',
  'gov.teams.microsoft.us',
  'teams.live.com',
  'teams.cloud.microsoft'
]

let pluginCtx = null
let secretCache = ''

function isAllowedTeamsUrl(url) {
  try {
    const parsed = new URL(url)
    if (parsed.protocol !== 'https:') return false
    if (parsed.username || parsed.password) return false
    const host = parsed.hostname.toLowerCase().replace(/\.$/, '')
    return TEAMS_WEB_HOSTS.some((name) => host === name || host.endsWith('.' + name))
  } catch {
    return false
  }
}

function openUrl(url) {
  if (!isAllowedTeamsUrl(url)) return
  const opener = pluginCtx?.os?.openExternal
  if (typeof opener === 'function') {
    void opener(url)
    return
  }
  window.open(url, '_blank', 'noopener,noreferrer')
}

function joinPath(root, ...parts) {
  const slash = String(root).includes('\\') ? '\\' : '/'
  const pieces = [String(root).replace(/[\\/]+$/, '')]
  for (const part of parts) {
    pieces.push(String(part).replace(/^[\\/]+|[\\/]+$/g, ''))
  }
  return pieces.join(slash)
}

async function readInstalledSecret() {
  const bridge = typeof window !== 'undefined' ? window.hermesDesktop : null
  if (!bridge || typeof bridge.readFileText !== 'function' || typeof bridge.agentPluginsRoot !== 'function') {
    return ''
  }
  let root = ''
  try {
    root = await bridge.agentPluginsRoot()
  } catch {
    return ''
  }
  if (!root) return ''
  try {
    const result = await bridge.readFileText(joinPath(root, ID, 'proxy.secret'))
    const text = String(result?.text || '').trim()
    if (!text || result?.truncated) return ''
    return text.split(/\s+/)[0]
  } catch {
    return ''
  }
}

async function ensureSecret(force) {
  if (secretCache && !force) return secretCache
  if (!force && pluginCtx?.storage) {
    const stored = pluginCtx.storage.get('proxySecret', '')
    if (stored) {
      secretCache = stored
      return stored
    }
  }
  const secret = await readInstalledSecret()
  if (secret) {
    secretCache = secret
    try {
      pluginCtx?.storage?.set('proxySecret', secret)
    } catch {
      // Persistence is a cache. The file remains the source of truth.
    }
  }
  return secret
}

async function rest(path) {
  if (!pluginCtx?.rest) {
    const err = new Error('backend-off')
    err.backend = false
    throw err
  }
  let secret = await ensureSecret(false)
  if (!secret) {
    // A rejected local call creates the secret file. It is not in the response.
    let probeError = null
    try {
      await pluginCtx.rest(path)
    } catch (err) {
      probeError = err
    }
    secret = await ensureSecret(true)
    if (!secret) {
      const probeMessage = String(probeError?.message || '')
      if (probeError && !/401|unauthorized/i.test(probeMessage)) {
        throw probeError
      }
      const err = new Error('proxy-secret-unavailable')
      throw err
    }
  }
  // The desktop REST door forwards the path, not custom headers. The secret
  // stays out of Graph; this process only talks to the local plugin API.
  const send = (value) => {
    const join = path.includes('?') ? '&' : '?'
    return pluginCtx.rest(`${path}${join}proxy_secret=${encodeURIComponent(value)}`)
  }
  try {
    return await send(secret)
  } catch (err) {
    secretCache = ''
    try {
      pluginCtx.storage?.remove('proxySecret')
    } catch {
      // Ignore a missing storage implementation.
    }
    const fresh = await readInstalledSecret()
    if (fresh && fresh !== secret) {
      secretCache = fresh
      try {
        pluginCtx.storage?.set('proxySecret', fresh)
      } catch {
        // See above.
      }
      return send(fresh)
    }
    throw err
  }
}

function TeamsPage() {
  const t = usePluginI18n(ID)
  const [teamId, setTeamId] = useState(null)
  const status = useQuery({
    queryKey: [ID, 'status'],
    queryFn: () => rest('/status'),
    staleTime: 30 * 1000
  })
  const teams = useQuery({
    queryKey: [ID, 'teams'],
    queryFn: () => rest('/teams'),
    enabled: Boolean(status.data?.ok),
    staleTime: 30 * 1000
  })
  const channels = useQuery({
    queryKey: [ID, 'channels', teamId],
    queryFn: () => rest(`/channels?team_id=${encodeURIComponent(teamId)}`),
    enabled: Boolean(teamId),
    staleTime: 30 * 1000
  })

  if (status.isLoading) {
    return jsx('div', {
      className: 'flex h-full items-center justify-center',
      children: jsx(GlyphSpinner, {})
    })
  }

  if (status.isError || status.data?.ok === false) {
    const message = String(status.error?.message || '')
    const backendOff = message === 'backend-off'
    const secretMissing = message === 'proxy-secret-unavailable'
    return jsx(ErrorState, {
      title: backendOff ? t('backendError') : t('authError'),
      description: backendOff
        ? t('backendHint')
        : secretMissing
          ? t('proxySecretHint')
          : status.data?.hint || status.data?.error || message,
      action: jsx(Button, {
        size: 'sm',
        onClick: () => {
          haptic('tap')
          void status.refetch()
        },
        children: t('retry')
      })
    })
  }

  const me = status.data?.me || {}
  const list = teams.data?.teams || []
  const selected = list.find((row) => row.id === teamId) || null
  const chList = channels.data?.channels || []

  return jsxs('div', {
    className: 'flex h-full min-h-0 flex-col',
    children: [
      jsxs('div', {
        className: 'flex items-center gap-2 px-3 py-2',
        children: [
          jsx(Codicon, { name: 'organization', className: 'text-(--ui-text-tertiary)' }),
          jsxs('div', {
            className: 'min-w-0 flex-1',
            children: [
              jsx('div', { className: 'truncate text-sm font-medium', children: t('paneTitle') }),
              jsx('div', {
                className: 'truncate text-[0.6875rem] text-(--ui-text-tertiary)',
                children: me.displayName || me.userPrincipalName || t('signedOut')
              })
            ]
          }),
          status.data?.teams_gateway_enabled
            ? jsx(Badge, { children: t('gatewayOn') })
            : jsx(Badge, { children: t('gatewayOff') }),
          jsx(Tip, {
            label: t('refreshTip'),
            children: jsx(Button, {
              size: 'sm',
              variant: 'ghost',
              onClick: () => {
                haptic('tap')
                void status.refetch()
                void teams.refetch()
                if (teamId) void channels.refetch()
              },
              children: t('refresh')
            })
          })
        ]
      }),
      jsx(Separator, {}),
      jsxs('div', {
        className: 'flex min-h-0 flex-1',
        children: [
          jsx(ScrollArea, {
            className: 'w-[240px] shrink-0 border-r border-(--ui-stroke-secondary)',
            children: jsx('div', {
              className: 'flex flex-col p-1',
              children: list.length
                ? list.map((row) =>
                    jsx(
                      'button',
                      {
                        type: 'button',
                        className: cn(
                          'flex w-full items-center gap-2 rounded px-2 py-1.5 text-left text-sm',
                          row.id === teamId
                            ? 'bg-(--chrome-action-hover) text-foreground'
                            : 'text-(--ui-text-secondary) hover:bg-(--chrome-action-hover)'
                        ),
                        onClick: () => {
                          haptic('tap')
                          setTeamId(row.id)
                        },
                        children: jsx('span', { className: 'truncate', children: row.displayName })
                      },
                      row.id
                    )
                  )
                : jsx('div', {
                    className: 'p-3 text-xs text-(--ui-text-tertiary)',
                    children: teams.isLoading ? t('loading') : t('noTeams')
                  })
            })
          }),
          jsx('div', {
            className: 'flex min-h-0 min-w-0 flex-1 flex-col',
            children: !selected
              ? jsx(EmptyState, { title: t('pickTeam'), description: t('pickHint') })
              : jsxs('div', {
                  className: 'flex h-full min-h-0 flex-col',
                  children: [
                    jsxs('div', {
                      className: 'flex items-center gap-2 px-3 py-2',
                      children: [
                        jsx('div', { className: 'min-w-0 flex-1 truncate text-sm font-medium', children: selected.displayName }),
                        selected.webUrl
                          ? jsx(Button, {
                              size: 'sm',
                              onClick: () => {
                                haptic('tap')
                                openUrl(selected.webUrl)
                              },
                              children: t('openTeam')
                            })
                          : null
                      ]
                    }),
                    jsx(Separator, {}),
                    channels.isLoading
                      ? jsx('div', {
                          className: 'flex flex-1 items-center justify-center',
                          children: jsx(GlyphSpinner, {})
                        })
                      : jsx(ScrollArea, {
                          className: 'flex-1',
                          children: jsx('div', {
                            className: 'flex flex-col gap-1 p-2',
                            children: chList.length
                              ? chList.map((ch) =>
                                  jsxs(
                                    'div',
                                    {
                                      className: 'flex items-center gap-2 rounded px-2 py-1.5',
                                      children: [
                                        jsx(Codicon, { name: 'comment-discussion', className: 'text-(--ui-text-quaternary)' }),
                                        jsx('span', { className: 'min-w-0 flex-1 truncate text-sm', children: ch.displayName }),
                                        ch.webUrl
                                          ? jsx(Button, {
                                              size: 'sm',
                                              variant: 'ghost',
                                              onClick: () => {
                                                haptic('tap')
                                                openUrl(ch.webUrl)
                                              },
                                              children: t('openChannel')
                                            })
                                          : null
                                      ]
                                    },
                                    ch.id
                                  )
                                )
                              : jsx('div', {
                                  className: 'p-3 text-xs text-(--ui-text-tertiary)',
                                  children: channels.data?.error || t('noChannels')
                                })
                          })
                        })
                  ]
                })
          })
        ]
      })
    ]
  })
}

export default {
  id: ID,
  name: 'Microsoft Teams',
  defaultEnabled: true,
  register(ctx) {
    pluginCtx = ctx
    ctx.i18n.register({
      en: {
        paneTitle: 'Teams',
        signedOut: 'Not signed in',
        gatewayOn: 'gateway on',
        gatewayOff: 'gateway off',
        refresh: 'Refresh',
        refreshTip: 'Reload Graph roster',
        retry: 'Retry',
        loading: 'Loading…',
        noTeams: 'No joined teams',
        pickTeam: 'Pick a team',
        pickHint: 'Joined teams from Microsoft Graph via Azure CLI.',
        openTeam: 'Open in Teams',
        openChannel: 'Open',
        noChannels: 'No channels',
        backendError: 'Teams backend is off',
        backendHint: 'Enable hermes-teams-inbox (`hermes plugins enable hermes-teams-inbox`) and reload Desktop.',
        authError: 'Graph sign-in needed',
        proxySecretHint:
          'The Teams proxy secret was not readable. It is created on first backend start at <HERMES_HOME>/plugins/hermes-teams-inbox/proxy.secret (mode 0600). Reload Desktop after az login.',
        chipTip: 'Microsoft Teams inbox'
      }
    })

    ctx.registerMany([
      {
        id: 'page',
        area: ROUTES_AREA,
        data: { path: ROUTE },
        render: () => jsx(TeamsPage, {})
      },
      {
        id: 'nav',
        area: SIDEBAR_NAV_AREA,
        data: { path: ROUTE, label: 'Teams', codicon: 'organization' }
      },
      {
        id: 'open',
        area: PALETTE_AREA,
        data: {
          id: 'hermes-teams-inbox.open',
          label: 'Open Microsoft Teams',
          keywords: ['teams', 'microsoft', 'graph', 'm365'],
          run: () => host.navigate(ROUTE)
        }
      }
    ])
  }
}
