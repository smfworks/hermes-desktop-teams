# Microsoft Teams for Hermes Desktop

There is **no official Hermes Desktop plugin** for Microsoft Teams. Hermes already ships a **gateway adapter** (`hermes-teams` — a bot people message) and a **meeting-summary pipeline**. This repo is the missing piece: a Desktop **inbox pane**.

SMF Works plugin id: `hermes-teams-inbox`.

## What it does

- Sidebar **Teams** page in Hermes Desktop
- Lists **joined teams** and **channels** from Microsoft Graph
- Opens the real Teams client/web URL for a channel
- Uses the machine's **Azure CLI** session (`az account get-access-token --resource https://graph.microsoft.com`) — no passwords in chat, no secrets in git
- Talks to Graph only through a **localhost proxy**. Every request needs a per-install secret the desktop client sends. See [Proxy auth](#proxy-auth)
- Chat/channel **message bodies** need extra Graph scopes (`ChannelMessage.Read.All`). Without them the API returns `needs_scope` instead of inventing a feed

## Install (this machine)

```text
HERMES_HOME\desktop-plugins\hermes-teams-inbox\plugin.js
HERMES_HOME\plugins\hermes-teams-inbox\   (Python API)
```

Then:

1. `hermes plugins enable hermes-teams-inbox` (or add it under `plugins.enabled` in `config.yaml`)
2. Desktop: Ctrl+K → **Reload desktop plugins**
3. Sidebar → **Teams**

Requires Azure CLI signed in (`az login`) with Graph access to `/me` and `/me/joinedTeams`.

## Proxy auth

The plugin API is not open just because Azure CLI is logged in. On first use it writes a secret to:

```text
<HERMES_HOME>/plugins/hermes-teams-inbox/proxy.secret
```

The file is mode `0600` and the directory is `0700`. Hermes Desktop reads that file (same Hermes home as the backend) and sends it as `proxy_secret`. You do not copy it by hand for a normal local install. Callers outside the app send it in the `X-Hermes-Teams-Proxy-Secret` header, `Authorization: Bearer`, or the `proxy_secret` query/body field.

A standalone server, if you run one, listens on **127.0.0.1:8765** only:

```text
python -m dashboard.plugin_api
```

Run that from the plugin directory (the repo root).

`HERMES_TEAMS_PROXY_PORT` changes the port. Binding any other address, or accepting non-loopback clients when Hermes mounts the routes, requires `HERMES_TEAMS_PROXY_ALLOW_REMOTE=1`. Leave that unset.

Only `/me`, joined teams, a team's channels, and a channel's messages are forwarded. `team_id` / `channel_id` are fully URL-encoded, and values containing `/` are rejected. `Open in Teams` only follows `https` links on Microsoft Teams hosts.

Tests (no live token): `pip install -r requirements-dev.txt && python -m pytest -q`.

## Not this plugin

- Enabling the Hermes **Teams gateway bot** (`platforms.teams`) — separate
- Meeting transcript pipeline — separate
