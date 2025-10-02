import os
import sys
import threading
import json
import subprocess
from pathlib import Path
import argparse

import PySimpleGUI as sg

BASE = Path(__file__).parent
LOG_DIR = BASE / 'logs'
LOG_FILE = LOG_DIR / 'jarvis.log'
CONV_HISTORY = LOG_DIR / 'conv-history.json'
JARVIS_SCRIPT = str(BASE / 'jarvis.py')
CLI_SCRIPT = str(BASE / 'cli_command.py')

proc = None


def get_python_exe() -> str:
    venv_py = BASE / '.venv' / 'Scripts' / 'python.exe'
    return str(venv_py) if venv_py.exists() else sys.executable


def start_jarvis():
    global proc
    if proc and proc.poll() is None:
        return
    python = get_python_exe()
    proc = subprocess.Popen([python, JARVIS_SCRIPT], cwd=str(BASE))


def stop_jarvis():
    global proc
    if proc and proc.poll() is None:
        try:
            proc.terminate()
            try:
                proc.wait(timeout=3)
            except Exception:
                proc.kill()
        except Exception:
            pass
    proc = None


def read_logs_tail(max_chars=4000):
    try:
        if not LOG_FILE.exists():
            return "(no logs yet)"
        text = LOG_FILE.read_text(encoding='utf-8', errors='ignore')
        return text[-max_chars:]
    except Exception:
        return "(failed to read logs)"


def read_history_tail():
    try:
        if not CONV_HISTORY.exists():
            return []
        data = json.loads(CONV_HISTORY.read_text(encoding='utf-8'))
        if not isinstance(data, list):
            return []
        return data[-50:]
    except Exception:
        return []


def run_oneoff_command(text: str):
    try:
        python = get_python_exe()
        # Run as a one-off so it speaks and exits
        subprocess.run([python, CLI_SCRIPT, text], cwd=str(BASE))
    except Exception:
        pass


def wake_once():
    try:
        python = get_python_exe()
        subprocess.run([python, CLI_SCRIPT, "--wake"], cwd=str(BASE))
    except Exception:
        pass


def load_contact_names(limit: int = 8):
    try:
        contacts_path = BASE / 'contacts.json'
        if not contacts_path.exists():
            return []
        data = json.loads(contacts_path.read_text(encoding='utf-8'))
        if not isinstance(data, dict):
            return []
        names = list(data.keys())
        # prioritize shorter names for buttons
        names.sort(key=lambda n: (len(n), n))
        return names[:limit]
    except Exception:
        return []


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--dark', action='store_true')
    args, _ = parser.parse_known_args()

    sg.theme('DarkBlue3' if args.dark else 'SystemDefault')
    contact_names = load_contact_names()
    if contact_names:
        contact_rows = [[sg.Text('Quick Call/Message')]] + [
            [sg.Button(f"Call: {name}"), sg.Button(f"Msg: {name}")]
            for name in contact_names
        ]
    else:
        contact_rows = [[sg.Text('No contacts found. Add entries to contacts.json')]]

    layout = [
        [sg.Text('Jarvis Control Panel', font=('Segoe UI', 14, 'bold'))],
        [sg.Button('Start'), sg.Button('Stop'), sg.Button('Wake'), sg.Button('Refresh Logs'), sg.Button('Refresh History')],
        [sg.Frame('Type Command', [[sg.Input(key='-CMD-', size=(70,1)), sg.Button('Run Command')]])],
        [
            sg.Frame('Quick Actions', [[
                sg.Button('Open WhatsApp'),
                sg.Button('Open YouTube'),
                sg.Button('Open Notepad'),
                sg.Button('Close Browser'),
                sg.Button('Call Mom'),
            ]])
        ],
        [
            sg.Frame('Protocols', [[
                sg.Button('Stealth Mode'),
                sg.Button('House Party'),
                sg.Button('Clean Slate'),
            ]]),
            sg.Frame('Utilities', [[
                sg.Button('Translate...'), sg.Button('Weather...'),
            ]])
        ],
        [
            sg.Frame('Contacts', contact_rows)
        ],
        [sg.Frame('Logs (tail)', [[sg.Multiline(read_logs_tail(), key='-LOGS-', size=(100, 20), autoscroll=True, disabled=True)]])],
        [sg.Frame('Conversation History (last 50)', [[sg.Multiline('', key='-HIST-', size=(100, 15), autoscroll=True, disabled=True)]])],
        [sg.Button('Open Log Folder'), sg.Button('Minimize to Tray'), sg.Button('Exit')]
    ]

    window = sg.Window('Jarvis GUI', layout, finalize=True)

    # initial populate
    window['-HIST-'].update(value=json.dumps(read_history_tail(), ensure_ascii=False, indent=2))

    while True:
        event, values = window.read(timeout=500)
        if event in (sg.WINDOW_CLOSED, 'Exit'):
            break
        if event == 'Start':
            start_jarvis()
        elif event == 'Stop':
            stop_jarvis()
        elif event == 'Wake':
            threading.Thread(target=wake_once, daemon=True).start()
        elif event == 'Run Command':
            text = (values.get('-CMD-') or '').strip()
            if text:
                window['-CMD-'].update('')
                threading.Thread(target=run_oneoff_command, args=(text,), daemon=True).start()
        elif event == 'Refresh Logs':
            window['-LOGS-'].update(value=read_logs_tail())
        elif event == 'Refresh History':
            window['-HIST-'].update(value=json.dumps(read_history_tail(), ensure_ascii=False, indent=2))
        elif event == 'Open Log Folder':
            os.startfile(str(LOG_DIR))
        elif event == 'Minimize to Tray':
            # Create a simple system tray with common actions
            window.hide()
            tray = sg.SystemTray(menu=['', ['Show', 'Start', 'Stop', 'Wake', 'Exit']], filename=None, tooltip='Jarvis')
            while True:
                ev = tray.read(timeout=1000)
                if ev in (None, ''):
                    continue
                if ev == 'Show':
                    window.un_hide()
                    tray.close()
                    break
                elif ev == 'Start':
                    start_jarvis()
                elif ev == 'Stop':
                    stop_jarvis()
                elif ev == 'Wake':
                    threading.Thread(target=wake_once, daemon=True).start()
                elif ev == 'Exit':
                    tray.close()
                    window.close()
                    return
        elif event == 'Open WhatsApp':
            threading.Thread(target=run_oneoff_command, args=("open whatsapp",), daemon=True).start()
        elif event == 'Open YouTube':
            threading.Thread(target=run_oneoff_command, args=("open youtube",), daemon=True).start()
        elif event == 'Open Notepad':
            threading.Thread(target=run_oneoff_command, args=("open notepad",), daemon=True).start()
        elif event == 'Close Browser':
            threading.Thread(target=run_oneoff_command, args=("close browser",), daemon=True).start()
        elif event == 'Call Mom':
            threading.Thread(target=run_oneoff_command, args=("call mom",), daemon=True).start()
        elif event == 'Stealth Mode':
            threading.Thread(target=run_oneoff_command, args=("engage stealth mode",), daemon=True).start()
        elif event == 'House Party':
            threading.Thread(target=run_oneoff_command, args=("house party protocol",), daemon=True).start()
        elif event == 'Clean Slate':
            threading.Thread(target=run_oneoff_command, args=("clean slate protocol",), daemon=True).start()
        elif event == 'Translate...':
            txt = sg.popup_get_text('Text to translate:', title='Translate')
            if txt:
                lang = sg.popup_get_text('Translate to (e.g., Hindi, English, Telugu):', title='Translate')
                if lang:
                    threading.Thread(target=run_oneoff_command, args=(f"translate {txt} to {lang}",), daemon=True).start()
        elif event == 'Weather...':
            city = sg.popup_get_text('City name (optional; blank for current location by IP):', title='Weather')
            threading.Thread(target=run_oneoff_command, args=((f"weather {city}" if city else "weather"),), daemon=True).start()
        else:
            # Dynamic contacts: match buttons like "Call: Name" or "Msg: Name"
            if isinstance(event, str) and event.startswith('Call: '):
                name = event.replace('Call: ', '', 1)
                threading.Thread(target=run_oneoff_command, args=(f"call {name}",), daemon=True).start()
            elif isinstance(event, str) and event.startswith('Msg: '):
                name = event.replace('Msg: ', '', 1)
                msg = sg.popup_get_text(f'Message to {name}:', title='Message')
                if msg is not None:
                    threading.Thread(target=run_oneoff_command, args=(f"message {name} {msg}",), daemon=True).start()


    window.close()


if __name__ == '__main__':
    main()
