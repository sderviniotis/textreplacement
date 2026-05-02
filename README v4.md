# Smart Text Replacer

**A privacy-first, open-source Windows text expansion utility.**

Replicates the macOS "Text Replacement" feature — type a short abbreviation and it expands to your full text, anywhere on Windows.

![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![Platform: Windows](https://img.shields.io/badge/Platform-Windows-0078d4.svg)
![Python: 3.10+](https://img.shields.io/badge/Python-3.10%2B-yellow.svg)

---

## Download

👉 **[Download SmartTextReplacer.exe from Releases](../../releases/latest)**

No Python required. Download and run.

> **Windows SmartScreen warning?** Click **More info → Run anyway**. This is expected for open-source utilities without a paid code signing certificate. The full source code is here for your review.

---

## Features

| Feature | Detail |
|---|---|
| **Word-boundary aware** | `ono` triggers on standalone `ono` — NOT inside `pro-bono` |
| **Password field safe** | Automatically disabled in password inputs and credential dialogs |
| **Dynamic variables** | `%DATE%`, `%TIME%`, `%DATETIME%`, `%DAY%`, `%MONTH%`, `%YEAR%`, `%CLIP%` |
| **App blocklist** | Block expansion in specific apps (Terminal, KeePass, etc.) |
| **Global hotkey** | `Ctrl+Alt+P` pauses/resumes without opening the app |
| **Encrypted storage** | Snippets encrypted with Fernet/AES-128 at rest |
| **Snippet groups** | Organise snippets into named categories |
| **Import / Export** | CSV import/export — compatible with spreadsheets |
| **Auto-backup** | Backup on every launch, last 10 kept |
| **System tray** | Runs silently in the background |
| **Windows startup** | Optional auto-launch on login (registry, no admin required) |
| **Update check** | Optional GitHub release check on startup |
| **Zero telemetry** | No network requests except optional update check |

---

## Dynamic Variables

Use these tokens anywhere in an expansion:

| Token | Inserts |
|---|---|
| `%DATE%` | Today's date — `12/04/2026` |
| `%TIME%` | Current time — `14:35` |
| `%DATETIME%` | Date + time — `12/04/2026 14:35` |
| `%DAY%` | Day name — `Sunday` |
| `%MONTH%` | Month name — `April` |
| `%YEAR%` | Year — `2026` |
| `%CLIP%` | Current clipboard contents |

**Example:** Set trigger `sig` → expansion `Kind regards,\nSteve Derviniotis\n%DATE%`

---

## Running from Source

```bash
pip install pynput pystray Pillow cryptography requests
pythonw text_replacer.py
```

Optional packages:
- `pystray` + `Pillow` — system tray icon
- `cryptography` — snippet encryption
- `requests` — update checking

---

## Building the EXE Yourself

```bash
pip install pyinstaller pynput pystray Pillow cryptography requests
pyinstaller --noconsole --onefile --name "SmartTextReplacer" text_replacer.py
```

Output: `dist/SmartTextReplacer.exe`

---

## Data & Privacy

All data is stored locally on your machine:

| File | Location |
|---|---|
| Snippets | `%APPDATA%\SmartTextReplacer\snippets.json` |
| Encryption key | `%APPDATA%\SmartTextReplacer\.key` |
| Backups | `%APPDATA%\SmartTextReplacer\backups\` |
| Log | `%APPDATA%\SmartTextReplacer\app.log` |

No data leaves your machine. No accounts. No cloud.

---

## License

MIT — free to use, modify, and distribute.

---

## Contributing

Issues and PRs welcome. If you find a bug or want to request a feature, open an issue.
