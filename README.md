# mcp-outlook-applescript

A [Model Context Protocol](https://modelcontextprotocol.io/) server for Microsoft Outlook on Mac. Access your mail, calendar, contacts, tasks, and notes through 53 MCP tools powered by AppleScript.

> **Security-hardened fork.** This fork adds two-phase approval for sending mail and for calendar writes, confines attachment/download file paths, neutralizes delimiter injection from untrusted content, and adds audit logging. See [Security & privacy](#security--privacy) — read it before pointing this at a work mailbox.

## Features

- **AppleScript backend** — works with classic Outlook for Mac, no network or authentication required
- **53 tools** — full coverage of mail, calendar, contacts, tasks, and notes
- **Approval system** — sending mail, destructive mail/folder operations (delete, move, junk, empty), and calendar writes (create, update, delete) require a two-step prepare/confirm flow
- **Path confinement** — downloaded attachments and outgoing-mail attachments are restricted to safe locations
- **Zero native dependencies** — pure TypeScript + AppleScript, no C++ compilation needed

### Available Tools

| Category | Tools |
|----------|-------|
| **Mail** | `list_folders`, `list_emails`, `search_emails`, `get_email`, `get_unread_count`, `prepare/confirm_send_email`, `list_attachments`, `download_attachment` |
| **Mail Organization** | `mark_email_read`, `mark_email_unread`, `set_email_flag`, `clear_email_flag`, `set_email_categories`, `prepare/confirm_delete_email`, `prepare/confirm_move_email`, `prepare/confirm_archive_email`, `prepare/confirm_junk_email`, `prepare/confirm_batch_*` |
| **Folders** | `create_folder`, `rename_folder`, `move_folder`, `prepare/confirm_delete_folder`, `prepare/confirm_empty_folder` |
| **Calendar** | `list_calendars`, `list_events`, `get_event`, `search_events`, `prepare/confirm_create_event`, `prepare/confirm_update_event`, `prepare/confirm_delete_event`, `respond_to_event` |
| **Contacts** | `list_contacts`, `search_contacts`, `get_contact` |
| **Tasks** | `list_tasks`, `search_tasks`, `get_task` |
| **Notes** | `list_notes`, `search_notes`, `get_note` |
| **Accounts** | `list_accounts` |

## Install

```bash
git clone https://github.com/MagnasiePro/mcp-outlook-applescript.git
cd mcp-outlook-applescript
npm install && npm run build && npm pack && npm install -g mcp-outlook-applescript-*.tgz
cd .. && rm -rf mcp-outlook-applescript
```

This installs `mcp-outlook-applescript` as a global command. The checkout can be deleted after install.

To update, re-run the commands above.

### Claude Code

```bash
claude mcp add outlook -- mcp-outlook-applescript
```

### Claude Desktop

Add to `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "mcp-outlook-applescript"
    }
  }
}
```

## Requirements

- macOS with Microsoft Outlook (classic) installed and running
- Node.js >= 20
- Automation permission for Outlook (System Settings > Privacy & Security > Automation)

## Security & privacy

This server gives an LLM broad read/write access to a mailbox. Understand the model before deploying it, especially on a corporate account.

**What the model can see.** Tools return full email bodies, attachments, contacts (including home addresses and phone numbers), calendar, notes, and tasks. All of that content flows to whatever model/provider your MCP client uses. Treat it as data egress: it may be logged or retained by that provider. Do not point this at a mailbox whose contents you cannot share with your model provider.

**Prompt-injection / "lethal trifecta".** The agent reads untrusted content (incoming email bodies) *and* can send mail — so a crafted email could try to instruct the agent to exfiltrate mailbox contents. Mitigations in this fork reduce but do not eliminate that risk:

- **Two-phase approval** guards `send_email` and calendar writes: the model must call `prepare_*` (which only previews) then `confirm_*`. This is a real safety boundary **only if your MCP client surfaces the confirm step to a human** — the same model can otherwise call both. Configure your client to require human approval for the `confirm_*` and destructive tools.
- **Path confinement.** Downloaded attachments are restricted to `OUTLOOK_MCP_DOWNLOAD_DIR` (default `~/Downloads`); outgoing-mail attachments cannot point at sensitive files (SSH/GPG/cloud credentials, private keys, `/etc`, dotfile config). Set `OUTLOOK_MCP_ATTACHMENT_DIR` to restrict outgoing attachments to a single directory.
- **Delimiter neutralization.** Untrusted content that could otherwise forge fields in tool output is neutralized on the detail-view and mail-listing read paths.
- **Audit log.** Send, delete, move, download, and calendar-write actions are logged as JSON to stderr for a basic forensic trail.

**Residual notes.** `respond_to_event` still sends an RSVP to the organizer without a separate confirm step. There is no authentication between the MCP client and this server: any process that can launch it inherits the Outlook Automation permission already granted. Approval tokens are single-use and expire after 5 minutes.

### Configuration

| Environment variable | Purpose | Default |
|----------------------|---------|---------|
| `OUTLOOK_MCP_DOWNLOAD_DIR` | Directory downloaded attachments must stay within | `~/Downloads` |
| `OUTLOOK_MCP_ATTACHMENT_DIR` | If set, outgoing attachments must stay within this directory (strict allowlist) | unset (denylist of sensitive paths) |

## Development

```bash
npm install
npm run build      # compile TypeScript → dist/
npm run typecheck   # type-check without emitting
npm test            # 145 unit tests
bash scripts/audit.sh  # static quality audit (build, security, package, functional)
```

## License

MIT — see [LICENSE](LICENSE).
