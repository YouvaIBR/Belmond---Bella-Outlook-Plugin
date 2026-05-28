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

## Mode 2 — Inject (function command + dialog)

**Ribbon button**: "Reply with Bella (Quick)"  
**Manifest action**: `ExecuteFunction` → `bellaReply`

### Flow

```
User clicks "Reply with Bella (Quick)"
  → No task pane opens
  → Notification "Bella is thinking…" appears in the email info bar
  → API call to n8n webhook (background)
  → Notification removed
  → Dialog opens (60% height, 40% width) with the draft editable
  → User edits the draft
  → User clicks "Copy" → content copied to clipboard
  → User clicks "Close" → dialog closes
```

### On error

A persistent error notification appears in the email info bar with a retry message.

---

## Why two modes?

The inject mode (function command) cannot call `displayReplyFormAsync` or
`displayNewMessageFormAsync` to open an Outlook reply window. This is a hard
platform limitation: function commands run in an ephemeral, UI-less runtime
that is destroyed as soon as `event.completed()` is called. Outlook does not
support shared runtimes, so there is no bridge between the function command
runtime and any compose window.

References:
- [GitHub OfficeDev/office-js #747](https://github.com/OfficeDev/office-js/issues/747) — `displayReplyFormAsync` silently fails in UI-less add-ins
- [Runtimes in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/runtimes) — Outlook does not support shared runtimes

`displayDialogAsync` is explicitly supported from function commands and is
the recommended workaround for any user-facing UI from a UI-less command.

---

## Architecture — inject mode runtime lifecycle

```
bellaReply(event)
  ├── notify() → "Bella is thinking…"
  ├── acquireToken() + callAgent()
  ├── localStorage.setItem("bella_draft", html)
  ├── displayDialogAsync("reply-dialog.html")
  │     ├── dialog loads → reads localStorage → renders draft
  │     ├── user edits / copies
  │     └── user clicks Close → messageParent("close")
  │           └── dialog.close() → event.completed()
  └── [on error] notify() → error message → event.completed()
```

`event.completed()` is deferred until the dialog closes. This keeps the
function command runtime alive for the duration of the dialog session.

---

## Files

| File | Role |
|---|---|
| `src/main.ts` | Task pane entry — auto-trigger, result panel, insert |
| `src/commands/commands.ts` | Function command — API call, notification, dialog |
| `src/dialog/dialog.ts` | Dialog logic — render draft, copy, close |
| `index.html` | Task pane HTML |
| `commands.html` | FunctionFile HTML (never displayed) |
| `reply-dialog.html` | Dialog HTML — editable draft UI |
| `manifest.xml` | Declares both buttons and FunctionFile |
