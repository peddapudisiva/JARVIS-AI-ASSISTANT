import sys
import argparse

import jarvis as core


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("command", nargs="?")
    parser.add_argument("--wake", action="store_true")
    args = parser.parse_args()

    core.load_env_if_available()
    cfg = core.load_json(core.CONFIG_PATH, {})
    custom_cmds = core.load_json(core.CUSTOM_CMDS_PATH, {})
    logger = core.init_logging()
    engine = core.init_tts(cfg)
    ai_model = core.init_ai(cfg)
    try:
        core.CURRENT_CFG.clear()
        if isinstance(cfg, dict):
            core.CURRENT_CFG.update(cfg)
    except Exception:
        pass

    if args.wake:
        try:
            if cfg.get("persona_enabled") and cfg.get("play_wake_chime"):
                import winsound
                winsound.Beep(1200, 90)
        except Exception:
            pass
        reply = (cfg.get("wake_reply") or "Yes?") if cfg.get("persona_enabled") else "Yes?"
        core.speak(engine, reply)
        return 0

    text = (args.command or "").strip()
    if not text:
        print("No command provided", file=sys.stderr)
        return 1

    intent, arg = core.parse_intent(text, custom_cmds)

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
        "call", "message", "email",
        # protocols and extras
        "protocol_stealth", "protocol_house_party", "protocol_clean_slate",
        "translate", "wiki", "weather",
        "unknown_open", "unknown_close", "prompt_open",
    }

    handled = False
    if cfg.get("ai_default_mode", True) and intent not in action_intents:
        handled = core.ai_or_search(engine, cfg, ai_model, text, logger)
    if not handled:
        core.execute_intent(engine, intent, arg)
    return 0


if __name__ == "__main__":
    sys.exit(main())
