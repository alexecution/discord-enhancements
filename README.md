# Discord Enhancements for NVDA

**Version 1.0.0**

This add-on enhances the experience of using the Discord Windows desktop application with the NVDA screen reader. It is inspired by [Doug Lee's JAWS scripts for Discord](https://dlee.org/discord) and brings a similar command-layer navigation system to NVDA.

---

## Key Concepts

### Browse Mode

Discord is an Electron (Chromium) application. By default, NVDA would present it in browse mode with virtual-cursor navigation, which can be confusing. This add-on **disables browse mode by default** so you interact with Discord's own controls directly.

You can still toggle browse mode at any time with `NVDA+Space` if you need virtual-cursor access.

### The Command Layer (Prefix Key)

Most features are accessed through a *command layer*. Press the **prefix key** (default: `[`), then press a follow-up key to run a command. For example, press `[` then `A` to hear the Active Now section.

- A short high-pitched tone plays when the layer is entered.
- A short low-pitched tone plays when a command is executed or the layer exits.
- Press `Escape` to cancel without running a command.
- Press the prefix key **twice quickly** to type it as a literal character.
- The layer auto-cancels after 5 seconds of inactivity.
- When you are typing in a message box the prefix key is passed through normally.

### Exploring Commands with Tab

While the command layer is active, press `Tab` or `Shift+Tab` to cycle through all available commands. NVDA will announce the key and description of each command. Press `Enter` to execute the last announced command. A wrap-around tone plays when you reach the end or beginning of the list.

### Window Title

When the Discord window is focused, NVDA reports the title with a *window* suffix (e.g. "My Server — Discord window"). Pressing `NVDA+T` gives an enhanced title that includes server name, channel, mute/deafen status, voice channel, badge counts, and on-screen alerts.

---

## Command Reference

### General

| Keys | Action |
|------|--------|
| `[` `A` | Announce the contents of the "Active Now" section (Friends page) |
| `[` `V` | Report which visible servers have active voice channels or live events |
| `[` `B` | List all buttons for activation (deduplicated, activates the chosen one) |

### Navigation

| Keys | Action |
|------|--------|
| `[` `E` | Move focus to the message input box |
| `[` `S` | Move focus to the server list |
| `[` `N` | Cycle focus among Discord areas (servers → channels → chat → members → user area) |
| `[` `U` | Move focus to the user area |

### Chat Message Navigation (Home-Row Keys)

| Keys | Action |
|------|--------|
| `[` `H` | Jump to the first available message |
| `[` `J` | Previous message |
| `[` `K` | Read current message (press twice quickly to spell) |
| `[` `L` | Next message |
| `[` `;` | Jump to the last available message |

### Chat Message Navigation (Arrow Keys)

| Keys | Action |
|------|--------|
| `[` `Home` | Jump to first message |
| `[` `Left Arrow` | Previous message |
| `[` `Numpad 5` | Read current message (press twice quickly to spell) |
| `[` `Right Arrow` | Next message |
| `[` `End` | Jump to last message |

### Shift Variants

| Keys | Action |
|------|--------|
| `[` `Shift+H` or `Shift+Home` | Jump to the "Unread" marker |
| `[` `Shift+K` or `Shift+Numpad 5` | Move real focus to the current message |
| `[` `Shift+P` | Open pinned messages |
| `[` `Shift+T` | Toggle the thread list |

### Recent Messages by Number

| Keys | Action |
|------|--------|
| `[` `1` | Read the most recent message |
| `[` `2` – `9` | Read the 2nd through 9th most recent message |
| `[` `0` | Read the 10th most recent message |

### Voice & Call Management

| Keys | Action |
|------|--------|
| `[` `D` | Disconnect from voice / stage channel |
| `[` `P` | Report ping / latency to the Discord server |

### Information

| Keys | Action |
|------|--------|
| `[` `T` | Announce who is typing and slow-mode status |
| `[` `W` | Announce channel info (topic, server name, DM details) |

### Other Gestures

| Keys | Action |
|------|--------|
| `NVDA+T` | Enhanced window title (server, channel, mute/deafen, voice, badges, alerts) |
| `NVDA+Space` | Toggle browse mode on/off (standard NVDA command) |

---

## Settings

Open **NVDA menu → Preferences → Settings → Discord Enhancements** to configure the add-on.

### Command Prefix Key

Change the prefix key used to enter the command layer. The default is `[`. You may use any key or key combination (e.g. `]`, `` ` ``). If the chosen key conflicts with an existing NVDA gesture, a warning dialog will appear.

> **Important:** If you change the prefix key and then uninstall the add-on, the key binding is automatically reset to the default to prevent orphaned bindings.

### Announce Incoming Chat Messages

When enabled (default), newly arriving messages in the focused channel are spoken automatically via live-region change events.

### Verbosity Level

Controls how much detail is reported for various UI elements:

- **Minimal (0)** — bare essentials only
- **Normal (1)** — recommended for most users (default)
- **Verbose (2)** — extra details for power users
- **Extra verbose (3)** — maximum detail, useful for debugging

---

## Enhanced Speech Output

The add-on automatically enriches the information NVDA speaks for various Discord elements:

- **Server list items** — voice activity and live-event indicators appended
- **Channel list items** — voice channel participants with mute/deafen status; "Limited" access notes
- **DM list items** — online status, friend nickname, group member count
- **Members list** — section heading announced (e.g. "Online — 5")
- **Chat messages** — structured reading of author, timestamp, and content
- **Forum posts** — title, tags, pinned status, age, reactions, message count, author
- **Polls** — question, options, vote counts and percentages
- **Inbox items** — reformatted with age prefixes and summaries

---

## Tips and Tricks

- **Quickly reach the Friends page:** Press `Ctrl+1` twice, then `[` `A` to hear the Active Now section.
- **Catch up with unread messages:** Press `[` `Shift+H` to jump to the unread marker, then use `[` `L` to read forward.
- **Discover commands:** Press `[` then `Tab` repeatedly to explore all available commands.
- **Spell a message:** Press `[` `K` twice quickly to spell the current message character by character.

---

## Troubleshooting

- **Commands not working:** Make sure focus is on the Discord window and you are not typing in a text field. The prefix key passes through to Discord when an edit field is focused.
- **"Element not found" messages:** Discord updates may change the UI structure. Check the NVDA log (`NVDA+F1`) for debug warnings about missing elements — these help developers update the add-on's element selectors.
- **Prefix key types a character instead:** You may be in a text field. Press `Escape` or `[` `E` to leave the text field first.

---

## Known Limitations

- Block quotes and code blocks are not specially indicated.
- Bullet points and numbered lists in messages may not be announced (use browse mode for full access).
- Braille output enhancements are not yet implemented.
- Some features depend on Discord's accessibility tree structure, which may change between versions.

---

## Acknowledgements

This add-on is inspired by [Doug Lee's JAWS scripts for Discord](https://dlee.org/discord). Thanks to the NVDA community for the screen reader and add-on framework.
