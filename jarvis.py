import os
import json
import subprocess
import webbrowser
import time
import threading
from datetime import datetime

try:
    import speech_recognition as sr
    import pyttsx3
    import pyautogui
    import pyperclip
    import keyboard
    import pygetwindow as gw
except Exception as e:
    print("Missing dependency:", e)
    print("Install required packages: SpeechRecognition pyaudio pyttsx3 pyautogui pyperclip keyboard pygetwindow")
    raise

# ---------- Config ----------
APP_MAP_FILE = "app_map.json"
VOICE_ENABLED = True      # if False, only text input is used
TTS_RATE = 160
WAKE_WORDS = ("jarvis", "hey jarvis", "ok jarvis")  # not enforced, but used in parsing if desired
LISTEN_TIMEOUT = 5        # seconds listening for speech
VOICE_PROMPT = "Listening..."


# ---------- Utilities ----------
def ensure_app_map():
    """Create default app_map.json if not present."""
    default = {
        "chrome": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        "notepad": "notepad.exe",
        "vscode": r"C:\Users\{}\AppData\Local\Programs\Microsoft VS Code\Code.exe".format(os.getlogin()),
        "calculator": "calc.exe",
        "edge": r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        "explorer": "explorer.exe",
        "python": r"C:\Users\{}\AppData\Local\Programs\Python\Python39-32\python.exe".format(os.getlogin())
    }
    if not os.path.exists(APP_MAP_FILE):
        with open(APP_MAP_FILE, "w", encoding="utf-8") as f:
            json.dump(default, f, indent=2)
        print(f"[Jarvis] Created default {APP_MAP_FILE}. Edit it to add your own app paths.")

def load_app_map():
    try:
        with open(APP_MAP_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print("[Jarvis] Failed to load app_map.json:", e)
        return {}

# ---------- TTS ----------
def init_tts():
    engine = pyttsx3.init()
    engine.setProperty('rate', TTS_RATE)
    return engine

tts_engine = init_tts()

speak_lock = threading.Lock()

def speak(text, block=False):
    """Say text and print to console."""
    print("Jarvis:", text)

    def _s():
        with speak_lock:
            try:
                tts_engine.say(text)
                tts_engine.runAndWait()
            except RuntimeError:
                pass  # safely ignore 'run loop already started'

    if block:
        _s()  # run synchronously
    else:
        # run in background thread so listen loop isn't blocked
        threading.Thread(target=_s, daemon=True).start()


# ---------- Speech recognition ----------
recognizer = sr.Recognizer()

def listen_once(timeout=LISTEN_TIMEOUT, phrase_time_limit=7):
    """Listen once and return recognized text or None."""
    try:
        with sr.Microphone() as source:
            recognizer.adjust_for_ambient_noise(source, duration=0.6)
            print(VOICE_PROMPT)
            audio = recognizer.listen(source, timeout=timeout, phrase_time_limit=phrase_time_limit)
        text = recognizer.recognize_google(audio)
        print("You:", text)
        return text.lower()
    except sr.WaitTimeoutError:
        print("[Jarvis] Listening timed out (no speech).")
        return None
    except sr.UnknownValueError:
        print("[Jarvis] Could not understand audio.")
        return None
    except Exception as e:
        print("[Jarvis] Microphone/listen error:", e)
        return None

# ---------- App & Window helpers ----------
def open_app(name, app_map):
    """Open an app by name using app_map or direct webbrowser for URLs."""
    name_lower = name.lower().strip()
    # If name looks like a URL:
    if name_lower.startswith("http") or "." in name_lower and " " not in name_lower:
        webbrowser.open(name)
        speak(f"Opening {name}")
        return True

    # check app_map
    if name_lower in app_map:
        path = app_map[name_lower]
        try:
            subprocess.Popen(path)
            speak(f"Opening {name_lower}")
            return True
        except Exception as e:
            speak(f"Failed to open {name_lower}: {e}")
            return False

    # try common names:
    try:
        subprocess.Popen(name_lower)
        speak(f"Opening {name_lower}")
        return True
    except Exception as e:
        speak(f"I don't know the path for {name}. Add it to {APP_MAP_FILE}.")
        return False

def activate_window_by_title(title_contains):
    """Try to activate first window that contains the given title text."""
    try:
        wins = gw.getWindowsWithTitle(title_contains)
        if not wins:
            # try case-insensitive find across all windows
            allw = gw.getAllTitles()
            for t in allw:
                if title_contains.lower() in (t or "").lower():
                    wins = gw.getWindowsWithTitle(t)
                    break
        if wins:
            w = wins[0]
            try:
                w.activate()
                time.sleep(0.3)
                return True
            except Exception:
                # try maximize/restore trick
                try:
                    w.maximize()
                    time.sleep(0.2)
                    w.restore()
                    time.sleep(0.2)
                    w.activate()
                except Exception:
                    pass
            return True
        return False
    except Exception as e:
        print("[Jarvis] Window activate error:", e)
        return False

# ---------- Command processing ----------
def do_search(query):
    speak(f"Searching Google for {query}")
    webbrowser.open(f"https://www.google.com/search?q={query.replace(' ', '+')}")

def copy_clipboard():
    try:
        text = pyperclip.paste()
        if text:
            speak("I have the clipboard content.")
            print("[Clipboard]:", text)
        else:
            speak("Clipboard is empty.")
        return text
    except Exception as e:
        speak("Failed to get clipboard.")
        print(e)
        return None

def paste_text_to_active_app(text):
    try:
        pyperclip.copy(text)
        # small delay to ensure clipboard is ready
        time.sleep(0.1)
        # simulate ctrl+v
        pyautogui.hotkey('ctrl', 'v')
        speak("Pasted.")
        return True
    except Exception as e:
        speak("Failed to paste.")
        print(e)
        return False

def type_text(text):
    try:
        pyautogui.write(text, interval=0.02)
        speak("Typed the text.")
        return True
    except Exception as e:
        speak("Typing failed.")
        print(e)
        return False

def run_command(cmd_text, app_map):
    """
    Parse and execute high-level commands.
    Returns True if command processed (even if failed), False otherwise (not recognized).
    """
    text = cmd_text.lower().strip()
    # small helper patterns
    if text in ("quit", "exit", "stop", "shutdown", "bye"):
        speak("Goodbye! Shutting down Jarvis.", block=True)
        raise SystemExit()

    if text.startswith("open "):
        name = text[5:].strip()
        # if phrase like "open youtube" or "open youtube.com"
        if "youtube" in name:
            webbrowser.open("https://www.youtube.com")
            speak("Opening YouTube")
            return True
        return open_app(name, app_map)

    if text.startswith("search ") or text.startswith("google "):
        # "search google for ..." or "search for ..."
        q = text
        for w in ("search ", "google "):
            if q.startswith(w): q = q[len(w):]
        if q.startswith("for "):
            q = q[4:]
        do_search(q)
        return True

    if text.startswith("copy"):
        # copy currently selected text (simulate ctrl+c)
        speak("Copying selection.")
        try:
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.1)
            copied = pyperclip.paste()
            print("[Copied]:", copied)
            speak("Copied to clipboard.")
        except Exception as e:
            speak("Failed to copy selection.")
            print(e)
        return True

    if text.startswith("paste into ") or text.startswith("paste to "):
        # example: "paste into vs code" or "paste into vscode"
        parts = text.split()
        # get app phrase after 'into' or 'to'
        if "into" in parts:
            app_name = " ".join(parts[parts.index("into")+1:])
        else:
            app_name = " ".join(parts[parts.index("to")+1:])
        speak(f"Switching to {app_name} and pasting.")
        ok = activate_window_by_title(app_name)
        time.sleep(0.3)
        if not ok:
            speak(f"Could not find window for {app_name}. I'll paste into the current window.")
        # paste from clipboard
        try:
            pyautogui.hotkey('ctrl', 'v')
            speak("Done.")
        except Exception as e:
            speak("Paste failed.")
            print(e)
        return True

    if text.startswith("paste") and "into" not in text:
        # simple paste
        speak("Pasting.")
        pyautogui.hotkey('ctrl', 'v')
        return True

    if text.startswith("type "):
        to_type = cmd_text[5:].strip()
        speak(f"Typing: {to_type}")
        type_text(to_type)
        return True

    if "time" in text:
        now = datetime.now().strftime("%I:%M %p")
        speak(f"The time is {now}")
        return True

    if text.startswith("run "):
        # example: run python manage.py runserver
        to_run = cmd_text[4:].strip()
        # if it seems to be a django server command, run in shell
        try:
            speak(f"Running: {to_run}")
            subprocess.Popen(to_run, shell=True)
            return True
        except Exception as e:
            speak("Failed to run the command.")
            print(e)
            return True

    # small chit-chat
    if any(x in text for x in ("hello", "hi", "hey")):
        speak("Hello! How can I assist you today?")
        return True
    if "how are you" in text:
        speak("I'm good, ready to help you.")
        return True

    return False  # not processed

# ---------- Main loop ----------
def main_loop():
    ensure_app_map()
    app_map = load_app_map()
    speak("Jarvis is ready. Say a command or type it here. Say 'quit' to exit.")

    while True:
        # Ask user whether they want to speak or type
        print("\nChoose input mode: [1] Voice  [2] Type  [3] Quick Text (press Enter)  [q] Quit")
        choice = input("Mode (1/2/q): ").strip().lower()
        if choice in ("q", "quit", "exit"):
            speak("Goodbye!", block=True)
            break

        user_text = None
        if choice == "1":
            if not VOICE_ENABLED:
                print("Voice disabled in config.")
                continue
            spoken = listen_once()
            if spoken:
                user_text = spoken
            else:
                print("No command recognized from voice.")
                continue
        elif choice == "2":
            user_text = input("Type your command: ").strip()
        else:
            # Quick mode: accept typed short command without menu
            user_text = input("Enter command (or press Enter to skip): ").strip()
            if not user_text:
                continue

        if not user_text:
            continue

        processed = run_command(user_text, app_map)
        if not processed:
            # fallback: if looks like a one-liner, try to interpret
            if user_text.startswith("http") or "." in user_text:
                webbrowser.open(user_text)
                speak("Opened link.")
            else:
                # default: search
                do_search(user_text)

if __name__ == "__main__":
    try:
        main_loop()
    except KeyboardInterrupt:
        print("\n[Jarvis] Interrupted by user. Exiting.")
    except SystemExit:
        print("\n[Jarvis] System exit requested.")
    except Exception as e:
        print("[Jarvis] Unexpected error:", e)
        raise
