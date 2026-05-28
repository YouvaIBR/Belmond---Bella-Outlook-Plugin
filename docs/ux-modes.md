# UX Modes — Bella Outlook Add-in

Two interaction modes are available, both visible as separate buttons in the Outlook ribbon.

---

## Mode 1 — Panel (task pane)

**Ribbon button**: "Reply with Bella"  
**Manifest action**: `ShowTaskpane`

### Flow

```
User clicks "Reply with Bella"
  → Task pane opens
  → "Bella is thinking…" shown immediately (auto-trigger)
  → API call to n8n webhook
  → Result rendered in panel (draft + enquiries + unanswered)
  → User edits the draft (contenteditable)
  → User clicks "Insert reply"
  → Outlook opens a reply form pre-filled with the draft
```

### On error

The generate button appears so the user can retry manually.

---

## Mode 2 — Inject (function command + event-based activation)

**Ribbon button**: "Reply with Bella (Quick)"  
**Manifest action**: `ExecuteFunction` → `bellaReply`

### Flow

```
User clicks "Reply with Bella (Quick)"
  → No task pane opens
  → Outlook reply form opens immediately with "Bella is thinking…" placeholder
  → API call to n8n webhook runs in background
  → OnNewMessageCompose event fires → handler polls sessionData for the draft
  → When draft is ready, body.setAsync() replaces the placeholder with real content
  → User edits and sends normally
```

### On error

A persistent error notification appears in the email info bar.

---

## Why two modes?

The inject mode uses `displayReplyAllForm()` called synchronously within the user gesture tick (before any `await`), which opens the reply window. The function command runtime continues running until `event.completed()` is called — this is deferred until the API call finishes so the draft can be written to `Office.sessionData`.

The `OnNewMessageCompose` event handler runs in the new compose window's context and reads `sessionData` to inject the draft via `body.setAsync()`. It polls up to 30 s (1 s intervals) in case the API call is still in flight when the compose window opens.

References:
- [Event-based activation](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/autolaunch) — OnNewMessageCompose
- [sessionData API](https://learn.microsoft.com/en-us/javascript/api/outlook/office.sessiondata) — per-item cross-runtime storage

---

## Architecture — inject mode runtime lifecycle

```
bellaReply(event)
  ├── sessionData.clearAsync()
  ├── displayReplyAllForm({ htmlBody: "Bella is thinking…" })   ← synchronous, user gesture
  └── runAgentAndStore(event, item)   ← async, runs in background
        ├── acquireToken() + callAgent()
        ├── sessionData.setAsync("bella_draft", html)
        └── event.completed()

OnNewMessageCompose fires in compose window
  └── injectDraft(event, item)
        ├── sessionData.getAsync("bella_draft")
        │     ├── found → body.setAsync(html) → done
        │     └── not found → pollForDraft() (up to 30 s)
        └── event.completed()
```

`event.completed()` in `bellaReply` is deferred until after `sessionData.setAsync()` so the data is available before the event handler polls.

---

## Files

| File | Role |
|---|---|
| `src/main.ts` | Task pane entry — auto-trigger, result panel, insert |
| `src/commands/commands.ts` | Function command — opens reply form, runs API, stores draft |
| `src/commands/launchEvents.ts` | OnNewMessageCompose handler — reads sessionData, injects draft |
| `index.html` | Task pane HTML |
| `commands.html` | FunctionFile HTML (never displayed) |
| `launchEvents.html` | Event handler HTML (never displayed) |
| `manifest.xml` | Declares both buttons, FunctionFile, and LaunchEvent |
