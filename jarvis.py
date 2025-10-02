import os
import time
import json
import threading
import re
import subprocess
import webbrowser
from datetime import datetime, timedelta
import logging
from logging.handlers import RotatingFileHandler
import winsound

import pyttsx3
import speech_recognition as sr
from urllib.parse import quote
import requests
from typing import List, Dict, Any
import smtplib
from email.mime.text import MIMEText
import sys

try:
    from dotenv import load_dotenv
except Exception:
    load_dotenv = None

try:
    from vosk import Model, KaldiRecognizer
except Exception:
    Model = None  # optional
    KaldiRecognizer = None

try:
    import keyboard  # global hotkey
except Exception:
    keyboard = None

try:
    import pywhatkit as kit
except Exception:
    kit = None

try:
    from comtypes import CLSCTX_ALL
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
except Exception:
    AudioUtilities = None
    IAudioEndpointVolume = None
    CLSCTX_ALL = None

try:
    import screen_brightness_control as sbc
except Exception:
    sbc = None

try:
    import google.generativeai as genai
except Exception:
    genai = None

try:
    import pyautogui  # optional input/window control
except Exception:
    pyautogui = None

try:
    import pygetwindow as gw  # optional precise window control
except Exception:
    gw = None


CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")
CUSTOM_CMDS_PATH = os.path.join(os.path.dirname(__file__), "custom_commands.json")
LOG_DIR = os.path.join(os.path.dirname(__file__), "logs")
LOG_FILE = os.path.join(LOG_DIR, "jarvis.log")
REMINDERS_PATH = os.path.join(os.path.dirname(__file__), "reminders.json")
CONV_HISTORY_PATH = os.path.join(LOG_DIR, "conv-history.json")
CONTACTS_PATH = os.path.join(os.path.dirname(__file__), "contacts.json")

# Global runtime config reference (set in main)
CURRENT_CFG: Dict[str, Any] = {}

WAKE_WORDS = ["jarvis", "jar viz", "jervis", "jar wish"]

# Whitelisted apps and actions for safety
WHITELISTED_APPS = {
    "notepad": "notepad",
    "calculator": "calc",
    "paint": "mspaint",
    "vscode": "code",
    "explorer": "explorer",
    # common browsers/apps
    "chrome": "chrome",
    "edge": "msedge",
    "firefox": "firefox",
    "brave": "brave",
    "opera": "opera",
    # popular desktop apps (may require being on PATH)
    "spotify": "spotify",
    "whatsapp": "whatsapp",
    "zoom": "zoom",
}

# Whitelisted process names for safe termination per app
WHITELISTED_APP_PROCESSES = {
    "notepad": ["notepad.exe"],
    "calculator": ["Calculator.exe", "ApplicationFrameHost.exe"],
    "paint": ["mspaint.exe"],
    "vscode": ["Code.exe"],
    "explorer": ["explorer.exe"],
    # browsers/apps
    "chrome": ["chrome.exe"],
    "edge": ["msedge.exe"],
    "firefox": ["firefox.exe"],
    "brave": ["brave.exe"],
    "opera": ["opera.exe", "opera_browser.exe", "opera_gx.exe"],
    # popular apps
    "spotify": ["Spotify.exe"],
    "whatsapp": ["WhatsApp.exe"],
    "zoom": ["Zoom.exe", "zoom.exe"],
}

# Common browser processes to close on "close browser"
BROWSER_PROCESSES = [
    "chrome.exe", "msedge.exe", "firefox.exe", "brave.exe", "opera.exe", "opera_browser.exe"
]

# App-specific launchers: try explicit paths/aliases when not on PATH
APP_LAUNCHERS: Dict[str, List[str]] = {
    "whatsapp": [
        os.path.expandvars(r"%LOCALAPPDATA%\WhatsApp\WhatsApp.exe"),
        os.path.expandvars(r"%USERPROFILE%\AppData\Local\WhatsApp\WhatsApp.exe"),
        "WhatsApp.exe", "whatsapp.exe"
    ],
    "spotify": [
        "Spotify.exe",
        os.path.expandvars(r"%APPDATA%\Spotify\Spotify.exe")
    ],
    "zoom": [
        os.path.expandvars(r"%APPDATA%\Zoom\bin\zoom.exe"),
        "Zoom.exe", "zoom.exe"
    ],
}

WHITELISTED_SITES = {
    "google": "https://www.google.com",
    "youtube": "https://www.youtube.com",
    "github": "https://github.com",
    "gmail": "https://mail.google.com",
    # commonly used sites
    "wikipedia": "https://www.wikipedia.org",
    "stackoverflow": "https://stackoverflow.com",
    "netflix": "https://www.netflix.com",
    "whatsapp": "https://web.whatsapp.com",
}


def load_json(path, default):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def init_logging():
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
        logger = logging.getLogger("jarvis")
        logger.setLevel(logging.INFO)
        if not logger.handlers:
            handler = RotatingFileHandler(LOG_FILE, maxBytes=512*1024, backupCount=3, encoding="utf-8")
            fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
            handler.setFormatter(fmt)
            logger.addHandler(handler)
        return logger
    except Exception:
        return logging.getLogger("jarvis_fallback")


def load_env_if_available():
    """Load variables from .env if python-dotenv is installed."""
    try:
        if load_dotenv:
            load_dotenv()
    except Exception:
        pass


SPEAK_LOCK = threading.Lock()

def safe_print(text: str):
    try:
        print(text)
    except UnicodeEncodeError:
        try:
            sys.stdout.reconfigure(encoding='utf-8', errors='ignore')  # Python 3.7+
        except Exception:
            pass
        try:
            print(text)
        except Exception:
            try:
                print(str(text).encode('utf-8', 'ignore').decode('utf-8', 'ignore'))
            except Exception:
                pass


def speak(engine: pyttsx3.Engine, text: str):
    # Serialize TTS to avoid 'run loop already started'
    with SPEAK_LOCK:
        try:
            engine.stop()
            engine.say(text)
            engine.runAndWait()
        except RuntimeError:
            # small retry once
            time.sleep(0.1)
            try:
                engine.stop()
                engine.say(text)
                engine.runAndWait()
            except Exception:
                pass


def init_tts(cfg):
    engine = pyttsx3.init()
    try:
        engine.setProperty("rate", cfg.get("response_rate", 180))
        engine.setProperty("volume", 1.0)
        pref = (cfg.get("voice_preference") or "").lower()
        voices = engine.getProperty("voices")
        if pref:
            for v in voices:
                if pref in v.name.lower():
                    engine.setProperty("voice", v.id)
                    break
    except Exception:
        pass
    return engine


def init_ai(cfg):
    provider = (cfg.get("ai_provider") or "").lower()
    if provider != "gemini" or not genai:
        return None
    api_key = cfg.get("google_api_key") or os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        return None
    try:
        genai.configure(api_key=api_key)
        model_name = cfg.get("gemini_model", "gemini-1.5-flash")
        model = genai.GenerativeModel(model_name)
        return model
    except Exception:
        return None


def ai_answer(model, question: str) -> str:
    if not model or not question:
        return ""
    try:
        resp = model.generate_content(question)
        text = getattr(resp, "text", None) or ""
        # Fallback: some SDK versions use .candidates
        if not text and getattr(resp, "candidates", None):
            for c in resp.candidates:
                parts = []
                for p in getattr(c, "content", {}).get("parts", []):
                    parts.append(p.get("text", ""))
                if parts:
                    text = "\n".join(parts)
                    break
        return text.strip()
    except Exception:
        return ""


def is_question(text: str) -> bool:
    if not text:
        return False
    t = text.strip().lower()
    if t.endswith('?'):
        return True
    starters = (
        'what', 'who', 'why', 'how', 'when', 'where', 'which',
        'explain', 'define', 'tell me', 'describe', 'compare', 'summarize'
    )
    return any(t.startswith(s) for s in starters)


def speak_ai_answer(engine: pyttsx3.Engine, cfg: dict, answer: str, logger: logging.Logger):
    if not answer:
        return
    try:
        if cfg.get("ai_print_full_answer", False):
            safe_print("\n=== AI Answer ===\n" + answer + "\n==================\n")
        max_chars = int(cfg.get("ai_tts_max_chars", 400))
        snippet = answer if len(answer) <= max_chars else answer[:max_chars]
        speak(engine, snippet)
    except Exception:
        # Fallback to basic speak
        speak(engine, answer[:400])


def speak_full_ai_answer(engine: pyttsx3.Engine, cfg: dict, answer: str):
    if not answer:
        speak(engine, "I don't have an answer to read yet")
        return
    chunk = int(cfg.get("ai_tts_max_chars", 500))
    chunk = max(200, min(chunk, 1200))
    for i in range(0, len(answer), chunk):
        speak(engine, answer[i:i+chunk])


def ai_or_search(engine: pyttsx3.Engine, cfg: dict, ai_model, query: str, logger: logging.Logger) -> bool:
    """Try AI answer; if unavailable/empty, open Google search instead. Returns True if handled."""
    if not query:
        return False
    ans = ai_answer(ai_model, query)
    if ans:
        speak_ai_answer(engine, cfg, ans, logger)
        # Optionally also open related web results even when AI answered
        try:
            if cfg.get("also_open_web_on_ai_answer", True):
                url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
                webbrowser.open(url)
        except Exception:
            pass
        return True
    # Fallback: open web search
    try:
        # honor config to disable web fallback
        if not cfg.get("web_fallback_on_ai_failure", True):
            speak(engine, "I don't have that answer right now.")
            return True
        url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
        webbrowser.open(url)
        speak(engine, f"Searching Google for {query}")
        return True
    except Exception:
        return False


def ai_route_intent(ai_model, text: str, logger: logging.Logger):
    """Use AI to classify a natural language command into one of our safe intents.
    Returns (intent, arg) or (None, None) if not confident.
    """
    if not ai_model or not text:
        return (None, None)
    try:
        system_prompt = (
            "You are a command router. Map the user's sentence to one of these intents: "
            "open_app, open_site, open_browser, search_web, search_youtube, time, date, greet, exit, "
            "volume, brightness, media, remind_in, remind_at, calc, convert, date_of_week, read_full_answer. "
            "Only choose intents that are obviously implied. If unsure, return intent 'none'.\n\n"
            "Return STRICT JSON with keys: intent, args. Where args depends on intent: \n"
            "- open_app: {target} (e.g., 'notepad', 'calculator', 'paint', 'vscode', 'explorer')\n"
            "- open_site: {target} (e.g., 'google', 'youtube', 'github', 'gmail')\n"
            "- open_browser: {}\n"
            "- search_web: {query}\n"
            "- search_youtube: {query}\n"
            "- time/date/greet/exit/read_full_answer: {}\n"
            "- volume: {direction} where direction in ['up','down','mute']\n"
            "- brightness: {direction} where direction in ['up','down']\n"
            "- media: {action} where action in ['play_pause','next','previous']\n"
            "- remind_in: {amount, unit, message} with unit in ['seconds','minutes','hours']\n"
            "- remind_at: {hour, minute, message} 24h integers\n"
            "- calc: {expr} using only digits +-*/().\n"
            "- convert: {value, src, dst} like 10, 'cm', 'inch'\n"
            "- date_of_week: {date} in YYYY-MM-DD.\n"
            "Respond with ONLY the JSON, no extra text."
        )

        prompt = f"User: {text}"
        resp = ai_model.generate_content([
            {"role": "user", "parts": [system_prompt + "\n" + prompt]}
        ])
        raw = getattr(resp, "text", "") or ""
        raw = raw.strip()
        # Some SDK versions may wrap in code fences; strip them
        if raw.startswith("```)" ):
            pass
        if raw.startswith("```json") and raw.endswith("```"):
            raw = raw[len("```json"): -3].strip()
        elif raw.startswith("```") and raw.endswith("```"):
            raw = raw[3:-3].strip()
        data = json.loads(raw)
        intent = (data.get("intent") or "").strip()
        if not intent or intent.lower() == "none":
            return (None, None)
        args = data.get("args")

        # Normalize to our executor's expectations
        if intent in ("open_app", "open_site"):
            target = (args.get("target") if isinstance(args, dict) else None) or None
            return (intent, (target or "").lower())
        if intent == "open_browser":
            return (intent, None)
        if intent in ("search_web", "search_youtube"):
            query = (args.get("query") if isinstance(args, dict) else None) or ""
            return (intent, query)
        if intent in ("time", "date", "greet", "exit", "read_full_answer"):
            return (intent, None)
        if intent == "volume":
            direction = (args.get("direction") if isinstance(args, dict) else None) or ""
            if direction not in ("up", "down", "mute"):
                return (None, None)
            return (intent, direction)
        if intent == "brightness":
            direction = (args.get("direction") if isinstance(args, dict) else None) or ""
            if direction not in ("up", "down"):
                return (None, None)
            return (intent, direction)
        if intent == "media":
            action = (args.get("action") if isinstance(args, dict) else None) or ""
            if action not in ("play_pause", "next", "previous"):
                return (None, None)
            return (intent, action)
        if intent == "remind_in":
            if isinstance(args, dict):
                amount = int(args.get("amount", 0))
                unit = str(args.get("unit", "")).lower()
                message = str(args.get("message", ""))
                if amount > 0 and unit in ("second", "seconds", "minute", "minutes", "hour", "hours") and message:
                    return (intent, (amount, unit, message))
        if intent == "remind_at":
            if isinstance(args, dict):
                hour = int(args.get("hour", -1))
                minute = int(args.get("minute", -1))
                message = str(args.get("message", ""))
                if 0 <= hour <= 23 and 0 <= minute <= 59 and message:
                    return (intent, (hour, minute, message))
        if intent == "calc":
            expr = (args.get("expr") if isinstance(args, dict) else None) or ""
            if re.fullmatch(r"[0-9\s\+\-\*\/\(\)\.]+", expr or ""):
                return (intent, expr)
        if intent == "convert":
            if isinstance(args, dict):
                try:
                    value = float(args.get("value"))
                except Exception:
                    value = None
                src = str(args.get("src", "")).lower()
                dst = str(args.get("dst", "")).lower()
                if value is not None and src and dst:
                    return (intent, (value, src, dst))
        if intent == "date_of_week":
            date_str = (args.get("date") if isinstance(args, dict) else None) or ""
            try:
                datetime.strptime(date_str, "%Y-%m-%d")
                return (intent, date_str)
            except Exception:
                return (None, None)
    except Exception:
        logger.info("AI routing failed or returned invalid JSON")
        return (None, None)

def recognize_speech(recognizer: sr.Recognizer, source: sr.AudioSource, cfg, timeout=5, phrase_time_limit=6) -> str:
    backend = (cfg.get("stt_backend") or "google").lower()
    language = cfg.get("language", "en-US")
    languages = cfg.get("languages") or []
    audio = recognizer.listen(source, timeout=timeout, phrase_time_limit=phrase_time_limit)
    try:
        if backend == "vosk" and Model and KaldiRecognizer:
            model_path = cfg.get("vosk_model_path")
            if not model_path or not os.path.isdir(model_path):
                return ""
            data = audio.get_wav_data(convert_rate=16000, convert_width=2)
            # Use Vosk recognizer
            model = Model(model_path)
            rec = KaldiRecognizer(model, 16000)
            rec.AcceptWaveform(data)
            result = json.loads(rec.Result())
            text = (result.get("text") or "").strip()
            return text
        else:
            # Google online recognizer with optional multi-language attempts
            # Try configured list first, then fall back to single language
            # By default, only try the primary language to avoid mis-detection.
            # If you explicitly enable try_all_languages in config, iterate the list.
            if bool(cfg.get("try_all_languages", False)):
                langs_to_try = [lang_code for lang_code in languages if isinstance(lang_code, str) and lang_code] or [language]
            else:
                langs_to_try = [language]
            for lang in langs_to_try:
                try:
                    text = recognizer.recognize_google(audio, language=lang)
                    return text.lower().strip()
                except sr.UnknownValueError:
                    continue  # try next language
            return ""
    except sr.WaitTimeoutError:
        return ""
    except sr.UnknownValueError:
        return ""
    except sr.RequestError:
        return ""


def contains_wake_word(text: str) -> bool:
    return any(w in text for w in WAKE_WORDS)


def parse_intent(command: str, custom_cmds: dict):
    c = command.lower().strip()

    # custom commands exact match
    if c in custom_cmds:
        act = custom_cmds[c]
        return (act.get("action"), act.get("target"))

    # open application
    if c == "open":
        return ("prompt_open", None)
    if c.startswith("open "):
        target = c.replace("open ", "", 1).strip()
        # prefer opening in browser if user asks explicitly (e.g., "open whatsapp in browser" or "open whatsapp web")
        if any(x in target for x in (" in browser", " on browser", " web")):
            site_hint = target
            for mark in (" in browser", " on browser", " web"):
                site_hint = site_hint.replace(mark, "")
            site_hint = site_hint.strip()
            # direct site key match first
            if site_hint in WHITELISTED_SITES:
                return ("open_site", site_hint)
            # special-case popular sites
            if site_hint in ("whatsapp", "whatsapp web"):
                return ("open_site", "whatsapp")
        if target in WHITELISTED_APPS:
            return ("open_app", target)
        if target in WHITELISTED_SITES:
            return ("open_site", target)
        # try patterns like open chrome/youtube/google
        if target in ("chrome", "browser"):
            return ("open_browser", None)
        # open any url or domain safely
        if target.startswith("http://") or target.startswith("https://"):
            return ("open_url", target)
        # simple domain match like example.com or sub.example.org
        if re.fullmatch(r"[a-z0-9.-]+\.[a-z]{2,}(\/.*)?", target):
            return ("open_url", f"https://{target}")
        return ("unknown_open", target)

    # close application/browser
    if c in ("close browser", "close the browser"):
        return ("close_browser", None)
    if c.startswith("close "):
        # normalize target after the verb
        target = c[6:].strip()
        # strip determiners like "the", "my"
        target = re.sub(r"^(the|my)\s+", "", target)
        # map common synonyms to browser close
        if target in ("browser", "chrome", "google", "edge", "firefox", "brave", "opera"):
            return ("close_browser", None)
        if target in WHITELISTED_APPS:
            return ("close_app", target)
        return ("unknown_close", target)

    # website shortcuts
    if c.startswith("go to ") or c.startswith("goto "):
        target = c.split(" ", 2)[-1].strip()
        if target in WHITELISTED_SITES:
            return ("open_site", target)
        # allow plain domains/urls
        if target.startswith("http://") or target.startswith("https://"):
            return ("open_url", target)
        if re.fullmatch(r"[a-z0-9.-]+\.[a-z]{2,}(\/.*)?", target):
            return ("open_url", f"https://{target}")
        return ("unknown_site", target)

    # search web
    if c.startswith("search ") or c.startswith("google "):
        query = c.split(" ", 1)[1].strip()
        return ("search_web", query)

    if c.startswith("youtube ") or c.startswith("search youtube "):
        q = c.split(" ", 1)[1].strip() if " " in c else ""
        return ("search_youtube", q)

    # input control: type text
    if c.startswith("type "):
        text = command[5:].strip()  # preserve original casing after 'type '
        if text:
            return ("type_text", text)

    # input control: press key or combo
    if c.startswith("press "):
        key = c[6:].strip()
        # normalize common names
        synonyms = {
            "enter": "enter", "return": "enter",
            "escape": "esc", "esc": "esc",
            "control": "ctrl", "ctrl": "ctrl",
            "alternate": "alt", "alt": "alt",
            "tab": "tab", "space": "space",
            "delete": "delete", "backspace": "backspace",
        }
        parts = [synonyms.get(p.strip(), p.strip()) for p in re.split(r"[+\-]", key) if p.strip()]
        if parts:
            return ("press_key", parts)

    # input control: scroll
    if c in ("scroll up", "scroll down", "scroll top", "scroll bottom"):
        direction = c.split(" ", 1)[1]
        return ("scroll", direction)

    # screenshot
    if c in ("screenshot", "take screenshot", "capture screen"):
        return ("screenshot", None)

    # time/date
    if "time" in c:
        return ("time", None)
    if "date" in c or "day" in c:
        return ("date", None)

    # volume
    if c in ("volume up", "increase volume"):
        return ("volume", "up")
    if c in ("volume down", "decrease volume"):
        return ("volume", "down")
    if c in ("mute", "unmute", "toggle mute"):
        return ("volume", "mute")

    # brightness
    if c in ("brightness up", "increase brightness"):
        return ("brightness", "up")
    if c in ("brightness down", "decrease brightness"):
        return ("brightness", "down")

    # media
    if c in ("play", "pause", "play pause", "resume"):
        return ("media", "play_pause")
    if c in ("next", "next track", "next song"):
        return ("media", "next")
    if c in ("previous", "previous track", "previous song"):
        return ("media", "previous")

    # reminders: "remind me in 10 minutes to stretch" or "remind me at 7:30 to ..."
    m = re.match(r"remind me in (\d+) (second|seconds|minute|minutes|hour|hours) to (.+)", c)
    if m:
        amount = int(m.group(1))
        unit = m.group(2)
        msg = m.group(3)
        return ("remind_in", (amount, unit, msg))
    m2 = re.match(r"remind me at (\d{1,2}):(\d{2}) to (.+)", c)
    if m2:
        hh = int(m2.group(1))
        mm = int(m2.group(2))
        msg = m2.group(3)
        return ("remind_at", (hh, mm, msg))

    # small talk
    if c in ("hello", "hi", "hey"):
        return ("greet", None)

    if c in ("stop", "exit", "quit", "bye"):
        return ("exit", None)

    # calculator: "calculate ..." or direct math "what is 2+2" / "2 + 2"
    if c.startswith("calculate "):
        expr = c.replace("calculate ", "", 1).strip()
        return ("calc", expr)
    mcalc = re.match(r"(what is |what's )?([0-9\s\+\-\*\/\(\)\.]+)$", c)
    if mcalc:
        expr = mcalc.group(2).strip()
        return ("calc", expr)

    # unit conversions: temperature, length, weight
    # examples: "convert 100 c to f", "convert 10 inches to cm", "convert 5 kg to lb"
    mconv = re.match(r"convert\s+([\d\.]+)\s*([a-z]+)\s+to\s+([a-z]+)", c)
    if mconv:
        val = float(mconv.group(1))
        src = mconv.group(2)
        dst = mconv.group(3)
        return ("convert", (val, src, dst))

    # date query: "what day is 2025-10-01"
    mdate = re.match(r"what (day|day of week) is (\d{4}-\d{2}-\d{2})", c)
    if mdate:
        return ("date_of_week", mdate.group(2))

    # read full last AI answer
    if c in ("read full answer", "read the answer", "read again", "repeat answer", "repeat the answer"):
        return ("read_full_answer", None)

    # persona protocols
    if c in ("engage stealth mode", "stealth mode", "enter stealth mode"):
        return ("protocol_stealth", None)
    if c in ("house party protocol", "initiate house party", "start house party"):
        return ("protocol_house_party", None)
    if c in ("clean slate protocol", "initiate clean slate", "clean slate"):
        return ("protocol_clean_slate", None)

    # communications: message/email/call
    # e.g., "message john i'm late", "send message to alice: meeting at 5"
    m_msg = re.match(r"(send\s+)?message\s+(to\s+)?([a-z\s]+?)[,:]?\s+(.*)$", c)
    if m_msg:
        name = m_msg.group(3).strip()
        text = m_msg.group(4).strip()
        if name and text:
            return ("message", (name, text))

    m_msg2 = re.match(r"(send\s+)?(a\s+)?message\s+to\s+([a-z\s]+)$", c)
    if m_msg2:
        name = m_msg2.group(3).strip()
        return ("message", (name, ""))

    # email commands
    m_mail = re.match(r"(send\s+)?email\s+(to\s+)?([a-z\s]+?)(?:\s+about|\s+regarding|\s+subject)?[,:]?\s*(.*)$", c)
    if m_mail:
        name = m_mail.group(3).strip()
        rest = (m_mail.group(4) or "").strip()
        return ("email", (name, rest))

    # call/dial commands
    m_call = re.match(r"(call|dial)\s+([a-z\s]+)$", c)
    if m_call:
        name = m_call.group(2).strip()
        return ("call", name)

    # weather: "weather", "weather in chennai", "what's the weather in mumbai"
    m_weather = re.match(r"(what'?s\s+the\s+)?(weather|temperature)(\s+in\s+(.+))?", c)
    if m_weather:
        loc = (m_weather.group(4) or "").strip()
        return ("weather", loc)

    # wikipedia summary: "who is ...", "what is ...", "tell me about ..."
    m_wiki = re.match(r"(who is|what is|tell me about)\s+(.+)$", c)
    if m_wiki:
        topic = m_wiki.group(2).strip()
        return ("wiki", topic)

    # timers: "set a timer for 10 minutes"
    m_timer = re.match(r"(set\s+)?(a\s+)?timer\s+for\s+(\d+)\s+(second|seconds|minute|minutes|hour|hours)", c)
    if m_timer:
        amount = int(m_timer.group(3))
        unit = m_timer.group(4)
        return ("remind_in", (amount, unit, "Timer finished"))

    # alarms: "set an alarm at 6", "set alarm for 6:30 am"
    m_alarm = re.match(r"(set\s+)?(an\s+)?alarm\s+(for|at)\s+(\d{1,2})(:(\d{2}))?\s*(am|pm)?", c)
    if m_alarm:
        hh = int(m_alarm.group(4))
        mm = int(m_alarm.group(6) or 0)
        ap = (m_alarm.group(7) or "").lower()
        if ap == "pm" and hh < 12:
            hh += 12
        if ap == "am" and hh == 12:
            hh = 0
        return ("remind_at", (hh, mm, "Alarm"))

    # translate: "translate <text> to <lang>"
    m_tr = re.match(r"translate\s+(.+?)\s+to\s+([a-zA-Z\-]+)$", c)
    if m_tr:
        text_to_tr = m_tr.group(1).strip()
        lang = m_tr.group(2).strip()
        return ("translate", (text_to_tr, lang))

    # news: "news", "headlines", "news about india", "tech news"
    m_news = re.match(r"(news|headlines)(\s+about\s+(.+))?$", c)
    if m_news:
        topic = (m_news.group(3) or "").strip()
        return ("news", topic)

    return ("unknown", c)


def contact_intent_from_text(text: str):
    """Lightweight fallback: infer call/message to a known contact by substring match.
    Returns (intent, arg) or (None, None).
    """
    try:
        c = (text or "").lower().strip()
        contacts = load_json(CONTACTS_PATH, {})
        names = list(contacts.keys())
        # call patterns
        if c.startswith("call ") or c.startswith("dial ") or " call " in c or " dial " in c:
            for name in names:
                if f"call {name}" in c or f"dial {name}" in c:
                    return ("call", name)
        # message patterns without body
        if c.startswith("message ") or " message " in c or c.startswith("send message"):
            for name in names:
                # with body
                for marker in (":", ",", " that ", " saying ", " say "):
                    key = f"message {name}{marker}"
                    if key in c:
                        body = c.split(key, 1)[1].strip()
                        return ("message", (name, body))
                # without body
                if f"message {name}" in c:
                    return ("message", (name, ""))
        return (None, None)
    except Exception:
        return (None, None)


def _volume_control(direction: str):
    if not AudioUtilities or not IAudioEndpointVolume:
        return False
    try:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = interface.QueryInterface(IAudioEndpointVolume)
        if direction == "mute":
            volume.SetMute(1, None) if not volume.GetMute() else volume.SetMute(0, None)
            return True
        # step by 0.1
        min_vol, max_vol, _ = volume.GetVolumeRange()
        current = volume.GetMasterVolumeLevelScalar()
        step = 0.1
        if direction == "up":
            volume.SetMasterVolumeLevelScalar(min(1.0, current + step), None)
        else:
            volume.SetMasterVolumeLevelScalar(max(0.0, current - step), None)
        return True
    except Exception:
        return False


def _brightness_control(direction: str):
    if not sbc:
        return False
    try:
        current = sbc.get_brightness(display=0)
        if isinstance(current, list):
            current = current[0]
        step = 10
        if direction == "up":
            sbc.set_brightness(min(100, current + step))
        else:
            sbc.set_brightness(max(10, current - step))
        return True
    except Exception:
        return False


def _media_key(action: str):
    # Use keyboard module to send media keys
    if not keyboard:
        return False
    try:
        if action == "play_pause":
            keyboard.send("play/pause media")
        elif action == "next":
            keyboard.send("next track")
        elif action == "previous":
            keyboard.send("previous track")
        return True
    except Exception:
        return False


def _terminate_processes_by_names(names: List[str]) -> bool:
    """Terminate processes by image names using taskkill. Returns True if any were killed.
    Only uses whitelisted names passed in.
    """
    if not names:
        return False
    any_killed = False
    for name in names:
        try:
            # /IM image name, /F force, /T terminate child processes
            subprocess.run(["taskkill", "/IM", name, "/F", "/T"], capture_output=True, text=True)
            any_killed = True
        except Exception:
            # continue trying others
            continue
    return any_killed


def schedule_reminder(engine: pyttsx3.Engine, when: datetime, message: str):
    # Persist this reminder
    try:
        os.makedirs(os.path.dirname(REMINDERS_PATH), exist_ok=True)
        reminders: List[Dict[str, Any]] = load_json(REMINDERS_PATH, [])
        reminders.append({
            "when": when.isoformat(),
            "message": message
        })
        with open(REMINDERS_PATH, "w", encoding="utf-8") as f:
            json.dump(reminders, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

    def worker():
        now = datetime.now()
        delay = max(0, (when - now).total_seconds())
        time.sleep(delay)
        speak(engine, f"Reminder: {message}")
        # Remove from persistence after firing
        try:
            reminders2: List[Dict[str, Any]] = load_json(REMINDERS_PATH, [])
            reminders2 = [r for r in reminders2 if not (
                r.get("message") == message and r.get("when") == when.isoformat()
            )]
            with open(REMINDERS_PATH, "w", encoding="utf-8") as f:
                json.dump(reminders2, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    t = threading.Thread(target=worker, daemon=True)
    t.start()


def restore_persistent_reminders(engine: pyttsx3.Engine):
    """Reload reminders from REMINDERS_PATH and reschedule future ones."""
    try:
        items: List[Dict[str, Any]] = load_json(REMINDERS_PATH, [])
        if not items:
            return
        now = datetime.now()
        remaining: List[Dict[str, Any]] = []
        for r in items:
            try:
                when = datetime.fromisoformat(r.get("when", ""))
                msg = str(r.get("message", "")).strip()
                if not msg:
                    continue
                if when > now:
                    # schedule again
                    schedule_reminder(engine, when, msg)
                    remaining.append({"when": when.isoformat(), "message": msg})
                # If past due, skip (do not fire late reminders on startup)
            except Exception:
                continue
        with open(REMINDERS_PATH, "w", encoding="utf-8") as f:
            json.dump(remaining, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def append_conv_history(entry: Dict[str, Any]):
    """Append a conversation entry to conv-history.json, keep last 200."""
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
        data: List[Dict[str, Any]] = []
        if os.path.exists(CONV_HISTORY_PATH):
            with open(CONV_HISTORY_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                if not isinstance(data, list):
                    data = []
        data.append(entry)
        # keep last 200
        data = data[-200:]
        with open(CONV_HISTORY_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def execute_intent(engine: pyttsx3.Engine, intent, arg):
    if intent == "open_app":
        app = WHITELISTED_APPS.get(arg)
        opened = False
        # Try explicit launchers first for apps that aren't on PATH
        try:
            from shutil import which
        except Exception:
            which = None  # type: ignore
        try:
            for candidate in APP_LAUNCHERS.get(arg, []):
                path = os.path.expandvars(candidate)
                # If it's an absolute file, prefer it; else try PATH lookup
                if os.path.isabs(path) and os.path.exists(path):
                    subprocess.Popen([path])
                    opened = True
                    break
                elif which and which(path):
                    subprocess.Popen([path])
                    opened = True
                    break
        except Exception:
            pass
        # Fallback to simple shell invocation
        if not opened and app:
            try:
                subprocess.Popen(app, shell=True)
                opened = True
            except Exception:
                opened = False
        if opened:
            speak(engine, f"Opening {arg}")
        else:
            # Do NOT fallback to web automatically; require explicit 'web' phrasing
            speak(engine, f"I couldn't open {arg}")

    elif intent == "open_browser":
        try:
            webbrowser.open("https://www.google.com")
            speak(engine, "Opening browser")
        except Exception:
            speak(engine, "I couldn't open the browser")

    elif intent == "open_site":
        url = WHITELISTED_SITES.get(arg)
        if url:
            webbrowser.open(url)
            speak(engine, f"Opening {arg}")
        else:
            speak(engine, f"Site {arg} is not allowed")

    elif intent == "prompt_open":
        speak(engine, "What should I open?")

    elif intent == "open_url":
        url = str(arg or "").strip()
        # final safety: only http(s) schemes
        if url.startswith("http://") or url.startswith("https://"):
            try:
                webbrowser.open(url)
                speak(engine, "Opening site")
            except Exception:
                speak(engine, "I couldn't open that site")
        else:
            speak(engine, "Invalid URL")

    elif intent == "close_browser":
        ok = _terminate_processes_by_names(BROWSER_PROCESSES)
        speak(engine, "Closed browser" if ok else "I couldn't close the browser")

    elif intent == "close_app":
        app = str(arg or "")
        # For safety, do not kill Windows shell (explorer)
        if app == "explorer":
            speak(engine, "Closing File Explorer is not supported for safety")
            return True
        procs = WHITELISTED_APP_PROCESSES.get(app, [])
        ok = _terminate_processes_by_names(procs)
        speak(engine, f"Closed {app}" if ok else f"I couldn't close {app}")

    elif intent == "unknown_open":
        # Clarify instead of routing to AI/search
        target = str(arg or "")
        if target:
            speak(engine, f"I can't open {target} yet. Say a known app or website.")
        else:
            speak(engine, "What should I open?")

    elif intent == "unknown_close":
        target = str(arg or "")
        if target:
            speak(engine, f"I can't close {target} yet. Say a known app.")
        else:
            speak(engine, "What should I close?")

    elif intent == "search_web":
        if arg:
            url = f"https://www.google.com/search?q={arg.replace(' ', '+')}"
            webbrowser.open(url)
            speak(engine, f"Searching Google for {arg}")
        else:
            speak(engine, "What should I search for?")

    elif intent == "search_youtube":
        q = arg or ""
        url = "https://www.youtube.com"
        if q:
            cfg2 = load_json(CONFIG_PATH, {})
            if (cfg2.get("youtube_play_top", True) and kit):
                try:
                    kit.playonyt(q)
                    speak(engine, f"Playing on YouTube: {q}")
                    return True
                except Exception:
                    pass
            url = f"https://www.youtube.com/results?search_query={q.replace(' ', '+')}"
            speak(engine, f"Searching YouTube for {q}")
        else:
            speak(engine, "Opening YouTube")
        webbrowser.open(url)

    elif intent == "time":
        now = datetime.now().strftime("%I:%M %p")
        speak(engine, f"It's {now}")

    elif intent == "date":
        today = datetime.now().strftime("%A, %B %d, %Y")
        speak(engine, f"Today is {today}")

    elif intent == "greet":
        speak(engine, "Hello, how can I help?")

    elif intent == "exit":
        speak(engine, "Goodbye")
        return False

    elif intent == "volume":
        ok = _volume_control(arg)
        speak(engine, "Done" if ok else "Volume control not available")

    elif intent == "brightness":
        ok = _brightness_control(arg)
        speak(engine, "Done" if ok else "Brightness control not available")

    elif intent == "media":
        ok = _media_key(arg)
        speak(engine, "Done" if ok else "Media control not available")

    elif intent == "remind_in":
        amount, unit, msg = arg
        mult = 1
        if unit.startswith("minute"):
            mult = 60
        elif unit.startswith("hour"):
            mult = 3600
        when = datetime.now() + timedelta(seconds=amount * mult)
        schedule_reminder(engine, when, msg)
        speak(engine, f"Reminder set in {amount} {unit}")

    elif intent == "remind_at":
        hh, mm, msg = arg
        now = datetime.now()
        when = now.replace(hour=hh, minute=mm, second=0, microsecond=0)
        if when < now:
            when += timedelta(days=1)
        schedule_reminder(engine, when, msg)
        speak(engine, f"Reminder set for {hh:02d}:{mm:02d}")

    elif intent == "calc":
        expr = str(arg)
        # only allow safe characters
        if not re.fullmatch(r"[0-9\s\+\-\*\/\(\)\.]+", expr):
            speak(engine, "I can only calculate basic arithmetic")
            return True
        try:
            result = eval(expr, {"__builtins__": {}}, {})
            speak(engine, f"The result is {result}")
        except Exception:
            speak(engine, "I couldn't compute that")
    
    elif intent == "protocol_stealth":
        # Mute and dim brightness
        _volume_control("mute")
        _brightness_control("down")
        speak(engine, "Stealth mode engaged")
    
    elif intent == "protocol_house_party":
        # Moderate volume, brighter screen, start music
        _brightness_control("up")
        try:
            if keyboard:
                keyboard.send("play/pause media")
        except Exception:
            pass
        speak(engine, "House Party Protocol activated")
    
    elif intent == "protocol_clean_slate":
        # Close whitelisted apps (except explorer), mute, and clear pending reminders file
        to_close = []
        for app, procs in WHITELISTED_APP_PROCESSES.items():
            if app == "explorer":
                continue
            to_close.extend(procs)
        _terminate_processes_by_names(list(set(to_close)))
        _volume_control("mute")
        try:
            with open(REMINDERS_PATH, "w", encoding="utf-8") as f:
                json.dump([], f)
        except Exception:
            pass
        speak(engine, "Clean Slate completed")

    elif intent == "win_minimize":
        handled = False
        try:
            if gw:
                w = gw.getActiveWindow()
                if w:
                    w.minimize()
                    handled = True
        except Exception:
            pass
        if not handled and pyautogui:
            try:
                pyautogui.hotkey("alt", "space")
                time.sleep(0.05)
                pyautogui.press("n")
                handled = True
            except Exception:
                pass
        speak(engine, "Minimized" if handled else "I couldn't minimize the window")

    elif intent == "win_restore":
        handled = False
        try:
            if gw:
                w = gw.getActiveWindow()
                if w:
                    w.restore()
                    handled = True
        except Exception:
            pass
        speak(engine, "Restored" if handled else "I couldn't restore the window")

    elif intent == "win_close":
        handled = False
        if pyautogui:
            try:
                pyautogui.hotkey("alt", "f4")
                handled = True
            except Exception:
                pass
        speak(engine, "Closed" if handled else "I couldn't close the window")

    elif intent == "win_switch":
        handled = False
        if pyautogui:
            try:
                pyautogui.hotkey("alt", "tab")
                handled = True
            except Exception:
                pass
        speak(engine, "Switching" if handled else "I couldn't switch windows")

    elif intent == "type_text":
        enabled = bool(CURRENT_CFG.get("input_control_enabled", False))
        if not enabled or not pyautogui:
            speak(engine, "Typing is disabled")
            return True
        try:
            pyautogui.typewrite(str(arg), interval=0.02)
            speak(engine, "Typed")
        except Exception:
            speak(engine, "I couldn't type")

    elif intent == "press_key":
        enabled = bool(CURRENT_CFG.get("input_control_enabled", False))
        if not enabled or not pyautogui:
            speak(engine, "Key press is disabled")
            return True
        keys = arg if isinstance(arg, list) else []
        try:
            if len(keys) >= 2:
                pyautogui.hotkey(*keys)
            elif len(keys) == 1:
                pyautogui.press(keys[0])
            else:
                raise ValueError("no keys")
            speak(engine, "Done")
        except Exception:
            speak(engine, "I couldn't press that")

    elif intent == "scroll":
        enabled = bool(CURRENT_CFG.get("input_control_enabled", False))
        if not enabled or not pyautogui:
            speak(engine, "Scrolling is disabled")
            return True
        step = int(CURRENT_CFG.get("scroll_step", 600))
        direction = str(arg or "down")
        try:
            if direction == "up":
                pyautogui.scroll(step)
            elif direction == "down":
                pyautogui.scroll(-step)
            elif direction == "top":
                pyautogui.keyDown("ctrl")
                pyautogui.press("home")
                pyautogui.keyUp("ctrl")
            elif direction == "bottom":
                pyautogui.keyDown("ctrl")
                pyautogui.press("end")
                pyautogui.keyUp("ctrl")
            speak(engine, "Scrolled")
        except Exception:
            speak(engine, "I couldn't scroll")

    elif intent == "screenshot":
        enabled = bool(CURRENT_CFG.get("input_control_enabled", False))
        if not enabled or not pyautogui:
            speak(engine, "Screenshot is disabled")
            return True
        try:
            ts = datetime.now().strftime("%Y%m%d-%H%M%S")
            out_dir = os.path.join(os.path.dirname(__file__), "screenshots")
            os.makedirs(out_dir, exist_ok=True)
            path = os.path.join(out_dir, f"screenshot-{ts}.png")
            img = pyautogui.screenshot()
            img.save(path)
            speak(engine, "Captured screenshot")
        except Exception:
            speak(engine, "I couldn't take a screenshot")
    elif intent == "convert":
        val, src, dst = arg
        src = src.lower()
        dst = dst.lower()
        try:
            converted = None
            # temperature
            if (src in ("c", "celsius") and dst in ("f", "fahrenheit")):
                converted = val * 9/5 + 32
            elif (src in ("f", "fahrenheit") and dst in ("c", "celsius")):
                converted = (val - 32) * 5/9
            # length
            elif (src in ("inch", "in", "inches") and dst in ("cm", "centimeter", "centimeters")):
                converted = val * 2.54
            elif (src in ("cm", "centimeter", "centimeters") and dst in ("inch", "in", "inches")):
                converted = val / 2.54
            elif (src in ("m", "meter", "meters") and dst in ("ft", "foot", "feet")):
                converted = val * 3.28084
            elif (src in ("ft", "foot", "feet") and dst in ("m", "meter", "meters")):
                converted = val / 3.28084
            # weight
            elif (src in ("kg", "kilogram", "kilograms") and dst in ("lb", "lbs", "pound", "pounds")):
                converted = val * 2.20462
            elif (src in ("lb", "lbs", "pound", "pounds") and dst in ("kg", "kilogram", "kilograms")):
                converted = val / 2.20462

            if converted is None:
                speak(engine, "I don't support that conversion yet")
            else:
                speak(engine, f"{val} {src} is {round(converted, 4)} {dst}")
        except Exception:
            speak(engine, "Conversion failed")

    elif intent == "date_of_week":
        try:
            d = datetime.strptime(arg, "%Y-%m-%d")
            speak(engine, d.strftime("That is a %A"))
        except Exception:
            speak(engine, "Invalid date format. Use YYYY-MM-DD")

    elif intent == "weather":
        loc = (arg or "").strip()
        try:
            url = f"https://wttr.in/{quote(loc)}?format=j1" if loc else "https://wttr.in?format=j1"
            r = requests.get(url, timeout=8)
            data = r.json()
            cur = data.get("current_condition", [{}])[0]
            temp_c = cur.get("temp_C")
            feels_c = cur.get("FeelsLikeC")
            desc = (cur.get("weatherDesc", [{}])[0].get("value") or "").lower()
            if temp_c is not None:
                if loc:
                    speak(engine, f"Weather in {loc}: {temp_c} degrees, feels like {feels_c}, {desc}")
                else:
                    speak(engine, f"Current weather: {temp_c} degrees, feels like {feels_c}, {desc}")
            else:
                raise ValueError("Missing temp")
        except Exception:
            try:
                webbrowser.open(f"https://www.google.com/search?q=weather+{quote(loc)}")
                speak(engine, "Opening weather")
            except Exception:
                speak(engine, "I couldn't get the weather")

    elif intent == "wiki":
        topic = str(arg)
        try:
            api = f"https://en.wikipedia.org/api/rest_v1/page/summary/{quote(topic)}"
            r = requests.get(api, timeout=8)
            if r.status_code == 200:
                summary = r.json().get("extract") or ""
                if summary:
                    speak(engine, summary[:500])
                else:
                    raise ValueError("No summary")
            else:
                raise ValueError("HTTP")
        except Exception:
            # Fallback to AI answer or open Wikipedia
            handled = ai_or_search(engine, {}, init_ai(load_json(CONFIG_PATH, {})), f"Provide a short summary about {topic}", logging.getLogger("jarvis"))
            if not handled:
                try:
                    webbrowser.open(f"https://en.wikipedia.org/wiki/{quote(topic)}")
                except Exception:
                    pass

    elif intent == "translate":
        text_to_tr, lang = arg
        cfg = load_json(CONFIG_PATH, {})
        model = init_ai(cfg)
        if model:
            prompt = f"Translate the following text into {lang}. Only return the translation.\n\nText: {text_to_tr}"
            ans = ai_answer(model, prompt)
            if ans:
                speak(engine, ans[:500])
                return True
        # Fallback: open Google Translate
        try:
            url = f"https://translate.google.com/?sl=auto&tl={quote(lang)}&text={quote(text_to_tr)}&op=translate"
            webbrowser.open(url)
            speak(engine, f"Opening translation to {lang}")
        except Exception:
            speak(engine, "I couldn't translate that")

    elif intent == "news":
        topic = (arg or "").strip()
        try:
            # Build Google News RSS URL (region India English by default)
            base = "https://news.google.com/rss?hl=en-IN&gl=IN&ceid=IN:en"
            url = base if not topic else f"https://news.google.com/rss/search?q={quote(topic)}&hl=en-IN&gl=IN&ceid=IN:en"
            r = requests.get(url, timeout=8)
            import xml.etree.ElementTree as ET
            root = ET.fromstring(r.text)
            titles = [item.findtext('title') for item in root.findall('.//item')]
            # Remove the feed title if present
            headlines = [t for t in titles if t and not t.lower().startswith('top stories')]
            if not headlines:
                raise ValueError('No headlines')
            top = headlines[:5]
            if topic:
                speak(engine, f"Top {len(top)} headlines about {topic}:")
            else:
                speak(engine, f"Top {len(top)} headlines:")
            for h in top:
                speak(engine, h[:200])
        except Exception:
            try:
                webbrowser.open("https://news.google.com/?hl=en-IN&gl=IN&ceid=IN:en")
                speak(engine, "Opening Google News")
            except Exception:
                speak(engine, "I couldn't get the news")

    elif intent == "message":
        # arg: (name, text)
        try:
            name, text = arg
            cfg = load_json(CONFIG_PATH, {})
            if not cfg.get("communications_enabled", True):
                speak(engine, "Messaging is disabled in settings")
                return True
            contacts = load_json(CONTACTS_PATH, {})
            key = name.lower().strip()
            info = contacts.get(key)
            if not info:
                speak(engine, f"I don't have contact info for {name}")
                return True
            channel = (cfg.get("default_message_channel") or "whatsapp").lower()
            if channel == "whatsapp":
                phone = (info.get("phone") or info.get("whatsapp") or "").replace(" ", "")
                if not phone:
                    speak(engine, f"{name} has no WhatsApp number saved")
                    return True
                if cfg.get("whatsapp_automation", False) and kit:
                    try:
                        # Send instantly; requires WhatsApp Web open and may bring browser to front
                        kit.sendwhatmsg_instantly(phone_no=phone, message=text or "")
                        speak(engine, f"Message sent to {name} on WhatsApp")
                        return True
                    except Exception:
                        pass
                # Fallback: open chat with prefilled text
                url = f"https://wa.me/{phone}?text={quote(text)}"
                webbrowser.open(url)
                speak(engine, f"Opening WhatsApp chat with {name}")
            elif channel == "email":
                email = info.get("email")
                if not email:
                    speak(engine, f"{name} has no email saved")
                    return True
                subject = ""
                body = text or ""
                cfg = load_json(CONFIG_PATH, {})
                if cfg.get("smtp_enabled", False) and cfg.get("smtp_host") and cfg.get("smtp_user"):
                    try:
                        msg = MIMEText(body, _charset="utf-8")
                        msg["Subject"] = subject
                        msg["From"] = cfg.get("smtp_user")
                        msg["To"] = email
                        server = smtplib.SMTP(cfg.get("smtp_host"), int(cfg.get("smtp_port", 587)))
                        if cfg.get("smtp_use_tls", True):
                            server.starttls()
                        smtp_pass = os.environ.get("SMTP_PASSWORD", "")
                        if smtp_pass:
                            server.login(cfg.get("smtp_user"), smtp_pass)
                        server.send_message(msg)
                        server.quit()
                        speak(engine, f"Email sent to {name}")
                        return True
                    except Exception:
                        # Fallback to mailto
                        pass
                from urllib.parse import quote
                url = f"mailto:{email}?subject={quote(subject)}&body={quote(body)}"
                webbrowser.open(url)
                speak(engine, f"Opening email to {name}")
            else:
                speak(engine, "Unsupported message channel")
        except Exception:
            speak(engine, "Failed to compose message")

    elif intent == "email":
        # arg: (name, text)
        try:
            name, text = arg
            contacts = load_json(CONTACTS_PATH, {})
            key = name.lower().strip()
            info = contacts.get(key)
            if not info or not info.get("email"):
                speak(engine, f"I don't have an email for {name}")
                return True
            cfg = load_json(CONFIG_PATH, {})
            subject = ""
            body = text or ""
            if cfg.get("smtp_enabled", False) and cfg.get("smtp_host") and cfg.get("smtp_user"):
                try:
                    msg = MIMEText(body, _charset="utf-8")
                    msg["Subject"] = subject
                    msg["From"] = cfg.get("smtp_user")
                    msg["To"] = info['email']
                    server = smtplib.SMTP(cfg.get("smtp_host"), int(cfg.get("smtp_port", 587)))
                    if cfg.get("smtp_use_tls", True):
                        server.starttls()
                    smtp_pass = os.environ.get("SMTP_PASSWORD", "")
                    if smtp_pass:
                        server.login(cfg.get("smtp_user"), smtp_pass)
                    server.send_message(msg)
                    server.quit()
                    speak(engine, f"Email sent to {name}")
                    return True
                except Exception:
                    pass
            from urllib.parse import quote
            url = f"mailto:{info['email']}?subject={quote(subject)}&body={quote(body)}"
            webbrowser.open(url)
            speak(engine, f"Opening email to {name}")
        except Exception:
            speak(engine, "Failed to compose email")

    elif intent == "call":
        # arg: name
        try:
            name = str(arg)
            cfg = load_json(CONFIG_PATH, {})
            contacts = load_json(CONTACTS_PATH, {})
            key = name.lower().strip()
            info = contacts.get(key)
            if not info or not info.get("phone"):
                speak(engine, f"I don't have a phone number for {name}")
                return True
            handler = (cfg.get("call_handler") or "tel").lower()
            number = info["phone"].replace(" ", "")
            if handler == "skype":
                url = f"skype:{number}?call"
                webbrowser.open(url)
                speak(engine, f"Trying to call {name} on Skype")
                return True
            if handler == "whatsapp":
                # Try WhatsApp desktop deep link, else web
                try:
                    webbrowser.open(f"whatsapp://send?phone={number}")
                    speak(engine, f"Opening WhatsApp call/chat for {name}")
                    return True
                except Exception:
                    pass
                try:
                    webbrowser.open(f"https://wa.me/{number}")
                    speak(engine, f"Opening WhatsApp Web chat for {name}")
                    return True
                except Exception:
                    pass
                speak(engine, "WhatsApp is not available")
                return True
            # Default: use tel: via Windows shell (more reliable than webbrowser for tel:)
            try:
                # Use start to invoke default dialer association
                subprocess.run(["cmd", "/c", "start", "", f"tel:{number}"], shell=True)
                speak(engine, f"Trying to call {name}")
                return True
            except Exception:
                pass
            # Fallback: try opening via webbrowser
            try:
                webbrowser.open(f"tel:{number}")
                speak(engine, f"Trying to call {name}")
            except Exception:
                speak(engine, "Failed to start call")
        except Exception:
            speak(engine, "Failed to start call")

    else:
        speak(engine, "I didn't understand that command")

    return True


def main():
    safe_print("Starting Jarvis Assistant (wake word 'jarvis' or hotkey)...")
    # Load .env first for secure config (e.g., GOOGLE_API_KEY)
    load_env_if_available()
    cfg = load_json(CONFIG_PATH, {})
    # Expose cfg to input-control handlers
    try:
        CURRENT_CFG.clear()
        if isinstance(cfg, dict):
            CURRENT_CFG.update(cfg)
    except Exception:
        pass
    custom_cmds = load_json(CUSTOM_CMDS_PATH, {})

    logger = init_logging()

    recognizer = sr.Recognizer()
    # Tunable STT parameters
    recognizer.energy_threshold = int(cfg.get("energy_threshold", 250))
    recognizer.dynamic_energy_threshold = bool(cfg.get("dynamic_energy_threshold", True))
    recognizer.pause_threshold = float(cfg.get("pause_threshold", 0.8))
    try:
        recognizer.non_speaking_duration = float(cfg.get("non_speaking_duration", 0.3))
    except Exception:
        pass

    engine = init_tts(cfg)
    # Restore reminders before we start listening
    restore_persistent_reminders(engine)
    ai_model = init_ai(cfg)

    # Microphone selection
    mic_index = None
    try:
        names = sr.Microphone.list_microphone_names()
        if names:
            safe_print("Available microphones (index: name):")
            for i, n in enumerate(names):
                safe_print(f"  {i}: {n}")
        cfg_mi = cfg.get("mic_index", None)
        if isinstance(cfg_mi, int):
            mic_index = cfg_mi
        elif isinstance(cfg_mi, str) and cfg_mi.strip().isdigit():
            mic_index = int(cfg_mi.strip())
        if mic_index is not None:
            safe_print(f"Using microphone index: {mic_index}")
    except Exception:
        pass

    hotkey = (cfg.get("hotkey") or "").lower()
    hotkey_read_full = (cfg.get("hotkey_read_full") or "").lower()
    wake_enabled = bool(cfg.get("wake_word_enabled", True))
    hotkey_triggered = {"flag": False}
    convo_window = int(cfg.get("conversation_window_seconds", 0))
    last_interaction = {"ts": 0.0}
    last_ai = {"text": ""}
    pending_command = {"text": None}
    last_empty_prompt = {"ts": 0.0}

    def on_hotkey():
        hotkey_triggered["flag"] = True

    if keyboard and (hotkey or hotkey_read_full):
        try:
            if hotkey:
                keyboard.add_hotkey(hotkey, on_hotkey)
                safe_print(f"Hotkey registered: {hotkey}")
            if hotkey_read_full:
                keyboard.add_hotkey(hotkey_read_full, lambda: speak_full_ai_answer(engine, cfg, last_ai["text"]))
                safe_print(f"Hotkey registered (read full): {hotkey_read_full}")
        except Exception:
            safe_print("Failed to register hotkey.")

    try:
        with sr.Microphone(device_index=mic_index) as source:
            safe_print("Calibrating for ambient noise...")
            calibrate_dur = float(cfg.get("ambient_noise_duration", 1.5))
            recognizer.adjust_for_ambient_noise(source, duration=calibrate_dur)
            speak(engine, "Jarvis online. Say my name or use the hotkey.")
            
            running = True
            while running:
                try:
                    now_ts = time.time()
                    in_conversation = convo_window > 0 and (now_ts - last_interaction["ts"]) <= convo_window

                    if not in_conversation:
                        if wake_enabled:
                            safe_print("Listening for wake word...")
                            wake_timeout = int(cfg.get("stt_timeout_wake", 12))
                            wake_phrase = int(cfg.get("stt_phrase_wake", 6))
                            text = recognize_speech(recognizer, source, cfg, timeout=wake_timeout, phrase_time_limit=wake_phrase)
                            if text:
                                safe_print(f"Heard: {text}")
                                logger.info(f"Heard wake loop: {text}")
                            if text and contains_wake_word(text):
                                # If user already asked the question with the wake word, use it directly.
                                tail = text
                                try:
                                    for ww in WAKE_WORDS:
                                        tail = tail.replace(ww, " ")
                                except Exception:
                                    pass
                                tail = tail.strip()
                                if tail:
                                    pending_command["text"] = tail
                                else:
                                    if bool(cfg.get("speak_prompt_on_wake", False)):
                                        try:
                                            if cfg.get("persona_enabled") and cfg.get("play_wake_chime"):
                                                winsound.Beep(1200, 90)
                                        except Exception:
                                            pass
                                        reply = (cfg.get("wake_reply") or "Yes?") if cfg.get("persona_enabled") else "Yes?"
                                        speak(engine, reply)
                            elif not hotkey_triggered["flag"]:
                                continue

                    if hotkey_triggered["flag"]:
                        hotkey_triggered["flag"] = False
                        if bool(cfg.get("speak_prompt_on_hotkey", False)):
                            try:
                                if cfg.get("persona_enabled") and cfg.get("play_wake_chime"):
                                    winsound.Beep(1200, 90)
                            except Exception:
                                pass
                            reply = (cfg.get("wake_reply") or "Yes?") if cfg.get("persona_enabled") else "Yes?"
                            speak(engine, reply)

                    safe_print("Waiting for command...")
                    cmd_timeout = int(cfg.get("stt_timeout_cmd", 12))
                    cmd_phrase = int(cfg.get("stt_phrase_cmd", 10))
                    # If we already have a command captured during wake, use it; else listen now
                    if pending_command["text"]:
                        command = pending_command["text"]
                        pending_command["text"] = None
                    else:
                        time.sleep(0.15)
                        command = recognize_speech(recognizer, source, cfg, timeout=cmd_timeout, phrase_time_limit=cmd_phrase)
                    if not command:
                        # Do not speak any prompt on empty recognition to avoid disturbance
                        logger.info("Empty command")
                        continue
                    safe_print(f"Command: {command}")
                    logger.info(f"Command: {command}")
                    intent, arg = parse_intent(command, custom_cmds)
                    try:
                        append_conv_history({
                            "ts": datetime.now().isoformat(),
                            "input": command,
                            "parsed_intent": intent,
                            "arg": arg
                        })
                    except Exception:
                        pass

                    action_intents = {
                        "open_app", "open_browser", "open_site",
                        "close_app", "close_browser",
                        "search_web", "search_youtube", "time", "date",
                        "volume", "brightness", "media",
                        "remind_in", "remind_at",
                        "calc", "convert", "date_of_week",
                        # window and input control
                        "win_minimize", "win_restore", "win_close", "win_switch",
                        "type_text", "press_key", "scroll", "screenshot",
                        "greet", "exit",
                        # communications
                        "call", "message", "email"
                    }

                    # If ai_default_mode is on, send everything to AI unless it matches explicit action intents
                    if cfg.get("ai_default_mode", False):
                        if intent not in action_intents:
                            handled = ai_or_search(engine, cfg, ai_model, command, logger)
                            if handled:
                                # Try to capture printed answer snippet from ai_or_search by re-calling ai_answer
                                ans = ai_answer(ai_model, command)
                                if ans:
                                    last_ai["text"] = ans
                                try:
                                    append_conv_history({
                                        "ts": datetime.now().isoformat(),
                                        "input": command,
                                        "ai_answer_len": len(ans or "")
                                    })
                                except Exception:
                                    pass
                                last_interaction["ts"] = time.time()
                                logger.info("Answered via AI/Search (default mode)")
                                continue

                    # Otherwise, if configured, route questions to AI by default
                    if cfg.get("ai_default_for_questions", False):
                        if intent not in action_intents and is_question(command):
                            handled = ai_or_search(engine, cfg, ai_model, command, logger)
                            if handled:
                                ans = ai_answer(ai_model, command)
                                if ans:
                                    last_ai["text"] = ans
                                try:
                                    append_conv_history({
                                        "ts": datetime.now().isoformat(),
                                        "input": command,
                                        "ai_answer_len": len(ans or "")
                                    })
                                except Exception:
                                    pass
                                last_interaction["ts"] = time.time()
                                logger.info("Answered via AI/Search (question default)")
                                continue

                    if intent == "unknown":
                        # Contact-aware quick match before AI routing
                        routed_intent, routed_arg = contact_intent_from_text(command)
                        if routed_intent:
                            running = execute_intent(engine, routed_intent, routed_arg)
                            last_interaction["ts"] = time.time()
                            logger.info(f"Executed via contact fallback: {routed_intent}")
                            continue
                        # Try AI action routing first if enabled
                        if cfg.get("ai_action_routing", True):
                            routed_intent, routed_arg = ai_route_intent(ai_model, command, logger)
                            if routed_intent:
                                running = execute_intent(engine, routed_intent, routed_arg)
                                last_interaction["ts"] = time.time()
                                logger.info(f"Executed via AI routing: {routed_intent}")
                                continue
                        # Fallback Q&A or search
                        handled = ai_or_search(engine, cfg, ai_model, command, logger)
                        if handled:
                            ans = ai_answer(ai_model, command)
                            if ans:
                                last_ai["text"] = ans
                            try:
                                append_conv_history({
                                    "ts": datetime.now().isoformat(),
                                    "input": command,
                                    "ai_answer_len": len(ans or "")
                                })
                            except Exception:
                                pass
                            last_interaction["ts"] = time.time()
                            logger.info("Answered via AI/Search (unknown)")
                            continue
                    # handle reading full answer locally
                    if intent == "read_full_answer":
                        speak_full_ai_answer(engine, cfg, last_ai["text"])
                        last_interaction["ts"] = time.time()
                        continue

                    running = execute_intent(engine, intent, arg)
                    try:
                        if cfg.get("persona_enabled") and cfg.get("play_completion_chime"):
                            winsound.Beep(900, 70)
                    except Exception:
                        pass
                    last_interaction["ts"] = time.time()
                    try:
                        append_conv_history({
                            "ts": datetime.now().isoformat(),
                            "input": command,
                            "executed_intent": intent
                        })
                    except Exception:
                        pass
                except sr.WaitTimeoutError:
                    pass
                except KeyboardInterrupt:
                    break
                except Exception as e:
                    safe_print(f"Error: {e}")
                    logger.exception("Unhandled error")
                    time.sleep(0.4)
    except KeyboardInterrupt:
        pass
    except OSError as e:
        safe_print(f"Microphone error: {e}")
        safe_print("Make sure a microphone is connected and not in use by another app.")
    finally:
        try:
            engine.stop()
        except Exception:
            pass
        if keyboard and (hotkey or hotkey_read_full):
            try:
                if hotkey:
                    keyboard.remove_hotkey(hotkey)
                if hotkey_read_full:
                    keyboard.remove_hotkey(hotkey_read_full)
            except Exception:
                pass
        safe_print("Jarvis stopped.")


if __name__ == "__main__":
    main()
