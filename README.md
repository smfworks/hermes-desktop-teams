# Microsoft Teams for Hermes Desktop

There is **no official Hermes Desktop plugin** for Microsoft Teams. Hermes already ships a **gateway adapter** (`hermes-teams` — a bot people message) and a **meeting-summary pipeline**. This repo is the missing piece: a Desktop **inbox pane**.

SMF Works plugin id: `hermes-teams-inbox`.

## What it does

- Sidebar **Teams** page in Hermes Desktop
- Lists **joined teams** and **channels** from Microsoft Graph
- Opens the real Teams client/web URL for a channel
- Uses the machine's **Azure CLI** session (`az account get-access-token --resource https://graph.microsoft.com`) — no passwords in chat, no secrets in git
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

## Not this plugin

- Enabling the Hermes **Teams gateway bot** (`platforms.teams`) — separate
- Meeting transcript pipeline — separate
