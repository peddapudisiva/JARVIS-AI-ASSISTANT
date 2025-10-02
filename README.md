# Jarvis – Minimal Local Voice Assistant (Windows)

![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0078D6?logo=windows)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
![Status](https://img.shields.io/badge/Status-Active-success.svg)
![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)

<!-- GitHub badges -->
[![GitHub stars](https://img.shields.io/github/stars/peddapudisiva/JARVIS-AI-ASSISTANT.svg?style=social)](https://github.com/peddapudisiva/JARVIS-AI-ASSISTANT)
[![GitHub issues](https://img.shields.io/github/issues/peddapudisiva/JARVIS-AI-ASSISTANT.svg)](https://github.com/peddapudisiva/JARVIS-AI-ASSISTANT/issues)
[![Build](https://github.com/peddapudisiva/JARVIS-AI-ASSISTANT/actions/workflows/ci.yml/badge.svg)](https://github.com/peddapudisiva/JARVIS-AI-ASSISTANT/actions)

A simple Python voice assistant named "Jarvis" that listens for the wake word "jarvis" and executes safe, whitelisted actions on Windows.
## Features
- Wake word detection: say "jarvis" to activate.
- Speech recognition using Google Web Speech API (no key required).
{{ ... }}
- Safe commands: open Notepad/Calculator/Paint, open popular sites, Google/Youtube search, tell time/date.
 - Offline STT option via Vosk (no internet required).
 - Global hotkey to trigger listening (default: Alt+J).
 - System controls: volume up/down/mute; screen brightness up/down (if supported).
 - Media controls: play/pause/next/previous.
 - Reminders: "remind me in 10 minutes to …" or "remind me at 7:30 to …".
 - Custom commands via `custom_commands.json`.
 - Multi-language recognition and TTS selection via `config.json`.
 - AI Q&A fallback using Google Gemini: ask any question and Jarvis will answer.

## Requirements
- Windows 10/11
- Python 3.9+ (64-bit recommended)
- A working microphone

## Setup
1. Create and activate a venv (recommended):
   ```powershell
   py -3 -m venv .venv
   .\\.venv\Scripts\Activate.ps1
   ```
2. Install dependencies:
   ```powershell
   pip install -r requirements.txt
   ```
   If `PyAudio` fails to install:
   - Install build tools: https://visualstudio.microsoft.com/visual-cpp-build-tools/
   - Or install a prebuilt wheel from https://www.lfd.uci.edu/~gohlke/pythonlibs/#pyaudio matching your Python version, e.g.:
     ```powershell
     pip install C:\Path\To\PyAudio‑0.2.14‑cp311‑cp311‑win_amd64.whl
     ```

## Optional: Offline speech recognition (Vosk)
Jarvis supports an offline STT backend using Vosk.

1. Download a Vosk model and extract it to `models/`, for example:
   - Small English: https://alphacephei.com/vosk/models (e.g., `vosk-model-small-en-us-0.15`)
2. Update `config.json`:
   ```json
   {
     "stt_backend": "vosk",
     "vosk_model_path": "models/vosk-model-small-en-us-0.15",
     "language": "en-US"
   }
   ```
3. Run Jarvis. If the model path is invalid, it will fall back silently.

## Optional: AI Q&A with Google Gemini
Jarvis can answer general questions using Google Gemini.

1. Get an API key: https://aistudio.google.com/app/apikey
2. Set the key either in `config.json` or as an environment variable:
   - In `config.json`:
     ```json
     {
       "ai_provider": "gemini",
       "gemini_model": "gemini-1.5-flash",
       "google_api_key": "YOUR_KEY_HERE"
     }
     ```
   - Or set an env var before running (PowerShell):
     ```powershell
     setx GOOGLE_API_KEY "YOUR_KEY_HERE"
     # restart terminal after setting
     ```
3. Usage: If Jarvis doesn't recognize a command, it will send your question to Gemini and speak back the answer.
   - Example: "jarvis" → "what is the tallest mountain in the world"
4. Privacy: The question is sent to Google if Gemini is enabled.

## Run
```powershell
python jarvis.py
```
- Wait for: "Jarvis online. Say my name to begin."
- Say: "jarvis". Jarvis will reply: "Yes?"
- Speak your command.
- Stop by saying: "stop" or press Ctrl+C in the terminal.

## Supported Commands (examples)
- Wake: "jarvis"
- Open app: "open notepad", "open calculator", "open paint"
- Open site: "open google", "go to youtube", "open browser"
- Search web: "search weather in chennai", "google python list comprehension"
- Search YouTube: "youtube lo-fi music"
- Time/Date: "what's the time", "what's the date"
- Greet/Exit: "hello", "stop" / "exit"

### System controls
- Volume: "volume up", "volume down", "mute"
- Brightness: "brightness up", "brightness down" (works on most laptops/monitors that expose brightness controls)

### Media controls
- "play" / "pause" / "play pause"
- "next", "previous"

### Reminders
- "remind me in 10 minutes to stretch"
- "remind me at 7:30 to join the meeting"

### Custom commands
Add phrases in `custom_commands.json` mapping to actions. Example:
```json
{
  "open vscode": { "action": "open_app", "target": "vscode" },
  "open explorer": { "action": "open_app", "target": "explorer" },
  "open github": { "action": "open_site", "target": "github" }
}
```
Supported actions: `open_app`, `open_site`, `media` (play_pause/next/previous), `volume` (up/down/mute).

## Safety
`jarvis.py` only executes whitelisted apps and sites defined in:
- `WHITELISTED_APPS`
- `WHITELISTED_SITES`
Edit those dictionaries to add more allowed targets.
- Recognition is inaccurate:
  - Reduce background noise; Jarvis auto-calibrates at startup.
  - Speak clearly after the wake word.
- TTS voice/rate:
  - Adjust in `init_tts()` by setting `rate`, `volume`, and a preferred `voice`.
- Global hotkey:
  - Default is `Alt+J`. Some systems require running the terminal as Administrator for global hotkeys.
  - Change in `config.json` under `hotkey` (empty string disables the hotkey).
  

## Customize
- Wake words: edit `WAKE_WORDS` in `jarvis.py`.
- Add intents: extend `parse_intent()` and `execute_intent()`.
- Add more commands safely by using whitelists and simple condition checks.
- Language and STT backend: edit `config.json` (`language`, `stt_backend`, `vosk_model_path`).
- Custom phrases: edit `custom_commands.json`.

## Notes
- The default recognizer uses an online API. If you need offline STT, consider Vosk.

## Quick Start

1. Create venv and install deps:
   ```powershell
   py -3 -m venv .venv
   .\.venv\Scripts\Activate.ps1
   pip install -r requirements.txt
   ```
2. Optional: Offline STT (Vosk)
   - Download `vosk-model-small-en-us-0.15` and extract to `models/`
   - Set `stt_backend` and `vosk_model_path` in `config.json` (see example: `config.example.json`)
3. Optional: Gemini AI Q&A
   - Get key from Google AI Studio and set in your environment or `config.json`
4. Run Jarvis:
   ```powershell
   python jarvis.py
   ```
5. Wake and speak commands:
   - Say "jarvis" → then your instruction
   - Or use hotkey (default Alt+J)

## Configuration Templates

- Copy `config.example.json` to `config.json` and adjust values
- Copy `contacts.example.json` to `contacts.json` and add your contacts

These files are git-ignored to protect your privacy.

## Contributing

- Fork the repo, create a feature branch, open a PR
- Keep commands safe by extending whitelists and intent parsing

## License

This project is released under the MIT License. See `LICENSE` for details.
