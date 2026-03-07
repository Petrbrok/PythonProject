import os
import json
import time
import random
import asyncio
import tempfile
import subprocess
import threading
import webbrowser
import queue
import re
import yaml
from datetime import datetime
from dotenv import load_dotenv
import edge_tts
import pygame

try:
    import win32com.client
    import pythoncom
    WSR_AVAILABLE = True
except ImportError:
    WSR_AVAILABLE = False

try:
    import vosk
    VOSK_AVAILABLE = True
except ImportError:
    VOSK_AVAILABLE = False

try:
    import sounddevice as sd
    SD_AVAILABLE = True
except ImportError:
    SD_AVAILABLE = False

try:
    import pvporcupine
    from pvrecorder import PvRecorder
    PORCUPINE_AVAILABLE = True
except ImportError:
    PORCUPINE_AVAILABLE = False

import speech_recognition

load_dotenv()

PICOVOICE_KEY   = os.getenv("PICOVOICE_KEY", "")
PPN_PATH        = os.path.join(os.path.dirname(os.path.abspath(__file__)), "lora_lora.ppn")
EDGE_VOICE      = os.getenv("EDGE_VOICE", "ru-RU-SvetlanaNeural")
WEATHER_API_KEY = os.getenv("WEATHER_API_KEY", "")

# ИИ отключён временно — не удалять
AI_ENABLED = False
try:
    from groq import Groq
    _groq_client = Groq(api_key=os.getenv("GROQ_API_KEY")) if os.getenv("GROQ_API_KEY") else None
except ImportError:
    _groq_client = None

pygame.mixer.init()
sr_engine = speech_recognition.Recognizer()
sr_engine.pause_threshold = 0.6

is_muted            = False
is_speaking         = False
last_activation     = 0.0
stop_speaking_event = threading.Event()

WINDOW_AFTER_AI      = 12
alarm_thread          = None
break_reminder_active = False

CONFIRM_PHRASES = ["Есть!", "Выполняю.", "Сделано.", "Готово.", "Принято."]
ERROR_PHRASES   = ["Не поняла.", "Повтори пожалуйста.", "Не расслышала."]
WAKE_PHRASES    = ["Слушаю.", "Да.", "Здесь."]

MUTE_TRIGGERS = (
    "замолчи", "молчи", "тихо", "пауза", "не слушай",
    "заткнись", "хватит", "подожди", "погоди", "тишина", "умолкни"
)
UNMUTE_TRIGGERS = (
    "размут", "включись", "продолжай", "проснись", "вернись", "активируйся"
)
STOP_TRIGGERS = (
    "завершить работу", "заверши работу", "выключись", "завершись",
    "закройся", "выход", "пока", "до свидания", "отключись", "выключи себя"
)

APP_PATHS = {
    "telegram":        r"C:\Users\%USERNAME%\AppData\Roaming\Telegram Desktop\Telegram.exe",
    "телеграм":        r"C:\Users\%USERNAME%\AppData\Roaming\Telegram Desktop\Telegram.exe",
    "discord":         r"C:\Users\%USERNAME%\AppData\Local\Discord\Update.exe",
    "дискорд":         r"C:\Users\%USERNAME%\AppData\Local\Discord\Update.exe",
    "spotify":         r"C:\Users\%USERNAME%\AppData\Roaming\Spotify\Spotify.exe",
    "спотифай":        r"C:\Users\%USERNAME%\AppData\Roaming\Spotify\Spotify.exe",
    "chrome":          r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    "хром":            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    "notepad":         "notepad.exe",
    "блокнот":         "notepad.exe",
    "calculator":      "calc.exe",
    "калькулятор":     "calc.exe",
    "paint":           "mspaint.exe",
    "диспетчер задач": "taskmgr.exe",
    "проводник":       "explorer.exe",
    "explorer":        "explorer.exe",
    "word":            r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    "excel":           r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
    "powerpoint":      r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
    "obs":             r"C:\Program Files\obs-studio\bin\64bit\obs64.exe",
    "steam":           r"C:\Program Files (x86)\Steam\steam.exe",
    "vscode":          r"C:\Users\%USERNAME%\AppData\Local\Programs\Microsoft VS Code\Code.exe",
    "код":             r"C:\Users\%USERNAME%\AppData\Local\Programs\Microsoft VS Code\Code.exe",
}

# ─── КОМАНДЫ ────────────────────────────────────────────────────────────────
# Добавляй новые команды сюда или в commands.yaml
# Формат yaml:
#   get_time:
#     - "который час"
#     - "сколько времени"

LOCAL_COMMANDS = {
    "который час":           "get_time",
    "сколько времени":       "get_time",
    "какое время":           "get_time",
    "время":                 "get_time",
    "какое сегодня число":   "get_date",
    "какая дата":            "get_date",
    "день недели":           "get_date",
    "увеличь громкость":     "volume_up",
    "громче":                "volume_up",
    "сделай громче":         "volume_up",
    "прибавь громкость":     "volume_up",
    "уменьши громкость":     "volume_down",
    "тише":                  "volume_down",
    "сделай тише":           "volume_down",
    "убавь громкость":       "volume_down",
    "выключи звук":          "sound_off",
    "без звука":             "sound_off",
    "включи звук":           "sound_on",
    "верни звук":            "sound_on",
    "увеличь яркость":       "brightness_up",
    "ярче":                  "brightness_up",
    "сделай ярче":           "brightness_up",
    "уменьши яркость":       "brightness_down",
    "темнее":                "brightness_down",
    "сделай темнее":         "brightness_down",
    "сверни окно":           "window_minimize",
    "сверни все":            "window_minimize",
    "сверни всё":            "window_minimize",
    "убери все окна":        "window_minimize",
    "разверни окно":         "window_maximize",
    "закрой окно":           "window_close",
    "переключи окно":        "switch_window",
    "альт таб":              "switch_window",
    "скопируй":              "clipboard_copy",
    "вставь":                "clipboard_paste",
    "что в буфере":          "clipboard_read",
    "скриншот":              "screenshot",
    "сделай скриншот":       "screenshot",
    "снимок экрана":         "screenshot",
    "заряд батареи":         "get_battery",
    "сколько заряда":        "get_battery",
    "батарея":               "get_battery",
    "загрузка процессора":   "get_cpu",
    "включи вайфай":         "wifi_on",
    "выключи вайфай":        "wifi_off",
    "открой браузер":        "open_browser",
    "закрой браузер":        "close_browser",
    "открой телеграм":       "open_app:телеграм",
    "запусти телеграм":      "open_app:телеграм",
    "открой дискорд":        "open_app:дискорд",
    "запусти дискорд":       "open_app:дискорд",
    "открой спотифай":       "open_app:спотифай",
    "открой хром":           "open_app:хром",
    "запусти хром":          "open_app:хром",
    "открой блокнот":        "open_app:блокнот",
    "открой калькулятор":    "open_app:калькулятор",
    "открой проводник":      "open_app:проводник",
    "открой стим":           "open_app:steam",
    "открой obs":            "open_app:obs",
    "открой код":            "open_app:код",
    "открой ворд":           "open_app:word",
    "открой эксель":         "open_app:excel",
    "включи музыку":         "play_music",
    "играй музыку":          "play_music",
    "выключи музыку":        "stop_music",
    "добавь задачу":         "create_task",
    "новая задача":          "create_task",
    "список задач":          "show_tasks",
    "мои задачи":            "show_tasks",
    "очисти задачи":         "clear_tasks",
    "выключи компьютер":     "shutdown",
    "перезагрузи":           "restart",
    "перезагрузка":          "restart",
    "спящий режим":          "sleep",
    "отмени выключение":     "cancel_shutdown",
    "открой загрузки":       "open_folder:загрузки",
    "открой документы":      "open_folder:документы",
    "открой рабочий стол":   "open_folder:рабочий стол",
    "открой музыку":         "open_folder:музыка",
    "открой видео":          "open_folder:видео",
    "включи перерывы":       "break_reminder_on",
    "выключи перерывы":      "break_reminder_off",
    "замолчи":               "mute",
    "молчи":                 "mute",
    "тихо":                  "mute",
    "размут":                "unmute",
    "включись":              "unmute",
}

_yaml_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "commands.yaml")
if os.path.exists(_yaml_path):
    try:
        _yaml_cmds = yaml.safe_load(open(_yaml_path, encoding="utf-8"))
        if _yaml_cmds:
            for _cmd, _phrases in _yaml_cmds.items():
                if isinstance(_phrases, list):
                    for _phrase in _phrases:
                        LOCAL_COMMANDS[_phrase.lower()] = _cmd
        print("  [yaml] Команды загружены")
    except Exception as e:
        print(f"  [!] commands.yaml: {e}")


# ─── WSR ГРАММАТИКА ─────────────────────────────────────────────────────────

class WSRListener:
    """Windows Speech Recognition с грамматикой — только известные фразы."""

    def __init__(self, phrases: list[str]):
        pythoncom.CoInitialize()
        self._recognizer = win32com.client.Dispatch("SAPI.SpInprocRecognizer")
        self._context    = self._recognizer.CreateRecoContext()
        self._grammar    = self._context.CreateGrammar(0)
        self._grammar.DictationSetState(0)  # отключаем диктовку

        rule = self._grammar.Rules.Add("commands",
               win32com.client.constants.SRATopLevel +
               win32com.client.constants.SRADynamic, 0)
        rule.Clear()
        for phrase in phrases:
            rule.InitialState.AddWordTransition(None, phrase)
        self._grammar.Rules.Commit()
        self._grammar.CmdSetRuleState("commands",
               win32com.client.constants.SGDSActive)

        self._result_queue = queue.Queue()

        # Подписка на события
        self._sink = win32com.client.WithEvents(
            self._context, _WSREventSink(self._result_queue)
        )

    def listen(self, timeout=8) -> str | None:
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                return self._result_queue.get(timeout=0.1)
            except queue.Empty:
                pass
        return None

    def close(self):
        try:
            self._grammar.CmdSetRuleState("commands",
                   win32com.client.constants.SGDSInactive)
        except Exception:
            pass


class _WSREventSink:
    def __init__(self, result_queue):
        self._q = result_queue

    def OnRecognition(self, StreamNumber, StreamPosition, RecognitionType, Result):
        try:
            res = win32com.client.Dispatch(Result)
            text = res.PhraseInfo.GetText().lower().strip()
            self._q.put(text)
        except Exception:
            pass


# ─── TTS ────────────────────────────────────────────────────────────────────

CACHED_PHRASES = {
    "ready":           "Готова к работе",
    "listening_again": "Снова слушаю",
    "confirm_0":       "Есть",
    "confirm_1":       "Выполняю",
    "confirm_2":       "Сделано",
    "confirm_3":       "Готово",
    "confirm_4":       "Принято",
    "unclear_0":       "Не поняла",
    "unclear_1":       "Повтори пожалуйста",
    "unclear_2":       "Не расслышала",
}
_phrase_cache: dict[str, str] = {}


def _generate_cache():
    cache_dir = os.path.join("sounds", "cache")
    wake_dir  = os.path.join("sounds", "wake")
    os.makedirs(cache_dir, exist_ok=True)
    os.makedirs(wake_dir, exist_ok=True)

    async def _gen(text, path):
        await edge_tts.Communicate(text, EDGE_VOICE, rate="+10%").save(path)

    for i, phrase in enumerate(WAKE_PHRASES):
        path = os.path.join(wake_dir, f"wake_{i}.mp3")
        if not os.path.exists(path):
            try:
                asyncio.run(_gen(phrase, path))
            except Exception:
                pass

    for key, phrase in CACHED_PHRASES.items():
        path = os.path.join(cache_dir, f"{key}.mp3")
        if not os.path.exists(path):
            try:
                asyncio.run(_gen(phrase, path))
            except Exception:
                pass
        if os.path.exists(path):
            _phrase_cache[key] = path


def _play_file(path: str):
    try:
        pygame.mixer.music.load(path)
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy():
            pygame.time.Clock().tick(10)
        pygame.mixer.music.unload()
    except Exception:
        pass


def play_cached(key: str):
    if key in _phrase_cache:
        print(f"  >> {CACHED_PHRASES.get(key, key)}")
        _play_file(_phrase_cache[key])
    elif key in CACHED_PHRASES:
        speak(CACHED_PHRASES[key])


def play_cached_random(prefix: str):
    keys = [k for k in _phrase_cache if k.startswith(prefix)]
    if keys:
        chosen = random.choice(keys)
        print(f"  >> {CACHED_PHRASES.get(chosen, chosen)}")
        _play_file(_phrase_cache[chosen])
    elif prefix == "confirm":
        speak(random.choice(CONFIRM_PHRASES))
    elif prefix == "unclear":
        speak(random.choice(ERROR_PHRASES))


def play_wake():
    wake_dir = os.path.join("sounds", "wake")
    files = [f for f in os.listdir(wake_dir) if f.endswith(".mp3")] if os.path.exists(wake_dir) else []
    if files:
        _play_file(os.path.join(wake_dir, random.choice(files)))
    else:
        speak(random.choice(WAKE_PHRASES))


def speak(text: str):
    global is_speaking
    if not text:
        return
    print(f"  >> {text}")

    def _worker():
        global is_speaking
        is_speaking = True
        stop_speaking_event.clear()
        tmp = None
        try:
            async def _synth():
                tts = edge_tts.Communicate(text, EDGE_VOICE)
                with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as f:
                    t = f.name
                await tts.save(t)
                return t
            tmp = asyncio.run(_synth())
            if tmp and not stop_speaking_event.is_set():
                pygame.mixer.music.load(tmp)
                pygame.mixer.music.play()
                while pygame.mixer.music.get_busy():
                    if stop_speaking_event.is_set():
                        pygame.mixer.music.stop()
                        break
                    pygame.time.Clock().tick(10)
                pygame.mixer.music.unload()
        except Exception as e:
            print(f"  [!] TTS: {e}")
        finally:
            is_speaking = False
            if tmp:
                try:
                    os.unlink(tmp)
                except Exception:
                    pass

    t = threading.Thread(target=_worker, daemon=True)
    t.start()
    t.join()


def stop_speech():
    global is_speaking
    stop_speaking_event.set()
    try:
        pygame.mixer.music.stop()
    except Exception:
        pass
    is_speaking = False


# ─── STT FALLBACK (Google) ──────────────────────────────────────────────────

def listen_google(timeout=8) -> str | None:
    try:
        with speech_recognition.Microphone() as source:
            audio = sr_engine.listen(source, phrase_time_limit=8, timeout=timeout)
        if is_speaking:
            stop_speech()
        text = sr_engine.recognize_google(audio, language="ru-RU").lower().strip()
        return text
    except speech_recognition.WaitTimeoutError:
        return None
    except speech_recognition.UnknownValueError:
        return None
    except Exception:
        return None


# ─── ИИ (отключён, не удалять) ──────────────────────────────────────────────

AI_SYSTEM_PROMPT = """Тебя зовут Лора — голосовой ассистент. Общаешься как друг — просто, с лёгким юмором. Отвечай на русском.
Простой вопрос — 1 предложение. Объяснение/творческое — столько сколько нужно.
Никогда не упоминай время и дату если не спросили. Время — только "Сейчас [время]".
Только текст. Никакого JSON."""


def ask_ai(query: str) -> str:
    if not AI_ENABLED or not _groq_client:
        return ""
    try:
        r = _groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            max_tokens=400,
            timeout=8,
            messages=[
                {"role": "system", "content": AI_SYSTEM_PROMPT},
                {"role": "user",   "content": f"[{datetime.now().strftime('%H:%M')}] {query}"},
            ],
        )
        return r.choices[0].message.content.strip()
    except Exception as e:
        return f"Ошибка: {e}"


# ─── КОМАНДЫ ────────────────────────────────────────────────────────────────

listen_fn = None


def get_time():
    speak(f"Сейчас {datetime.now().strftime('%H:%M')}.")
    return True

def get_date():
    months = ["января","февраля","марта","апреля","мая","июня",
              "июля","августа","сентября","октября","ноября","декабря"]
    n = datetime.now()
    speak(f"Сегодня {n.day} {months[n.month-1]} {n.year}, "
          f"{['понедельник','вторник','среда','четверг','пятница','суббота','воскресенье'][n.weekday()]}.")
    return True

def screenshot():
    try:
        import pyautogui
        pyautogui.screenshot(f"screenshot_{int(time.time())}.png")
        return True
    except Exception:
        return False

def clipboard_read():
    try:
        import pyperclip
        t = pyperclip.paste()
        speak(f"В буфере: {t[:200]}" if t else "Буфер пуст.")
        return True
    except Exception:
        return False

def _resolve_app(name: str) -> str:
    return APP_PATHS.get(name.lower(), name).replace("%USERNAME%", os.environ.get("USERNAME", ""))

def open_app(name: str):
    p = _resolve_app(name)
    try:
        subprocess.Popen(p)
        return True
    except Exception:
        try:
            subprocess.Popen(p, shell=True)
            return True
        except Exception:
            return False

def close_app(name: str):
    pm = {
        "telegram":"telegram","телеграм":"telegram","discord":"discord","дискорд":"discord",
        "spotify":"spotify","chrome":"chrome","хром":"chrome","notepad":"notepad",
        "word":"winword","excel":"excel","obs":"obs64","steam":"steam","vscode":"code","код":"code",
    }
    try:
        subprocess.run(["powershell.exe", f"Stop-Process -Name {pm.get(name.lower(),name)} -ErrorAction SilentlyContinue"], check=False)
        return True
    except Exception:
        return False

def open_browser():
    speak("Какой сайт?")
    site = listen_fn()
    if not site:
        return False
    url = f"https://{site}" if "." in site and " " not in site else f"https://www.google.com/search?q={site.replace(' ','+')}"
    webbrowser.open(url)
    return True

def close_browser():
    for b in ("chrome","firefox","msedge","opera","brave"):
        try:
            subprocess.run(["powershell.exe", f"Stop-Process -Name {b} -ErrorAction SilentlyContinue"], check=False)
        except Exception:
            pass
    return True

def _vol():
    try:
        from ctypes import POINTER, cast
        from comtypes import CLSCTX_ALL
        from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
        i = AudioUtilities.GetSpeakers().Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        return cast(i, POINTER(IAudioEndpointVolume))
    except Exception:
        return None

def sound_off():
    v = _vol()
    if v:
        v.SetMute(1, None)
    else:
        os.system("powershell.exe (new-object -com wscript.shell).SendKeys([char]173)")
    return True

def sound_on():
    v = _vol()
    if v:
        v.SetMute(0, None)
    else:
        os.system("powershell.exe (new-object -com wscript.shell).SendKeys([char]173)")
    return True

def volume_up():
    v = _vol()
    if v:
        v.SetMasterVolumeLevelScalar(min(1.0, v.GetMasterVolumeLevelScalar() + 0.1), None)
    else:
        for _ in range(5):
            os.system("powershell.exe (new-object -com wscript.shell).SendKeys([char]175)")
    return True

def volume_down():
    v = _vol()
    if v:
        v.SetMasterVolumeLevelScalar(max(0.0, v.GetMasterVolumeLevelScalar() - 0.1), None)
    else:
        for _ in range(5):
            os.system("powershell.exe (new-object -com wscript.shell).SendKeys([char]174)")
    return True

def brightness_up():
    try:
        import screen_brightness_control as sbc
        sbc.set_brightness(min(100, sbc.get_brightness()[0] + 20))
        return True
    except Exception:
        os.system("powershell (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,[math]::Min(100,((Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightness).CurrentBrightness)+20))")
        return True

def brightness_down():
    try:
        import screen_brightness_control as sbc
        sbc.set_brightness(max(0, sbc.get_brightness()[0] - 20))
        return True
    except Exception:
        os.system("powershell (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,[math]::Max(0,((Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightness).CurrentBrightness)-20))")
        return True

def switch_window():
    import pyautogui; pyautogui.hotkey("alt","tab"); return True
def window_minimize():
    import pyautogui; pyautogui.hotkey("win","d"); return True
def window_maximize():
    import pyautogui; pyautogui.hotkey("win","up"); return True
def window_close():
    import pyautogui; pyautogui.hotkey("alt","f4"); return True
def clipboard_copy():
    import pyautogui; pyautogui.hotkey("ctrl","c"); return True
def clipboard_paste():
    import pyautogui; pyautogui.hotkey("ctrl","v"); return True

def clipboard_write():
    speak("Что записать?")
    t = listen_fn()
    if not t:
        return False
    try:
        import pyperclip; pyperclip.copy(t); return True
    except Exception:
        return False

def get_battery():
    try:
        import psutil
        b = psutil.sensors_battery()
        if b is None:
            speak("Батарея не найдена.")
        else:
            speak(f"Заряд {int(b.percent)}%, {'заряжается' if b.power_plugged else 'на батарее'}.")
        return True
    except Exception:
        return False

def get_cpu():
    try:
        import psutil
        speak(f"Процессор {psutil.cpu_percent(interval=1)}%, память {psutil.virtual_memory().percent}%.")
        return True
    except Exception:
        return False

def wifi_off():
    os.system('netsh interface set interface "Wi-Fi" disabled'); return True
def wifi_on():
    os.system('netsh interface set interface "Wi-Fi" enabled'); return True

def get_weather():
    speak("Какой город?")
    city = listen_fn()
    if not city:
        return False
    if not WEATHER_API_KEY:
        speak("Добавь WEATHER_API_KEY в .env")
        return False
    try:
        import urllib.request, urllib.parse
        url = f"https://api.openweathermap.org/data/2.5/weather?q={urllib.parse.quote(city)}&appid={WEATHER_API_KEY}&units=metric&lang=ru"
        with urllib.request.urlopen(url, timeout=5) as r:
            d = json.loads(r.read().decode())
        speak(f"В {city}: {d['weather'][0]['description']}, {round(d['main']['temp'])}°, ощущается {round(d['main']['feels_like'])}°.")
        return True
    except Exception:
        speak("Не удалось получить погоду.")
        return False

def create_task():
    speak("Что добавить?")
    t = listen_fn()
    if not t:
        return False
    with open("список дел.txt", "a", encoding="utf-8") as f:
        f.write(f"✅ {t}\n")
    speak(f"Добавила: {t}")
    return True

def show_tasks():
    try:
        with open("список дел.txt", encoding="utf-8") as f:
            tasks = f.read().strip()
        if not tasks:
            speak("Список пуст.")
        else:
            lines = tasks.splitlines()
            speak(f"Задач {len(lines)}: " + "; ".join(l.replace("✅ ","") for l in lines[:5]))
        return True
    except Exception:
        return False

def clear_tasks():
    open("список дел.txt", "w", encoding="utf-8").close()
    return True

def play_music():
    try:
        files = [f for f in os.listdir("music") if f.endswith((".mp3",".wav",".flac"))]
        if not files:
            speak("В папке music нет файлов.")
            return False
        f = os.path.join("music", random.choice(files))
        os.startfile(f)
        speak(f"Включаю {os.path.splitext(os.path.basename(f))[0]}")
        return True
    except Exception:
        return False

def stop_music():
    for p in ("wmplayer","vlc","spotify","groove","musicbee"):
        try:
            subprocess.run(["powershell.exe", f"Stop-Process -Name {p} -ErrorAction SilentlyContinue"], check=False)
        except Exception:
            pass
    return True

def set_timer(seconds: int):
    def _t():
        time.sleep(seconds)
        speak("Таймер сработал!")
    threading.Thread(target=_t, daemon=True).start()
    m, s = divmod(seconds, 60)
    speak(f"Таймер на {m} мин." if m else f"Таймер на {s} сек.")
    return True

def set_alarm(hour: int, minute: int):
    global alarm_thread
    def _a():
        while True:
            n = datetime.now()
            if n.hour == hour and n.minute == minute:
                speak(f"Будильник! {hour:02d}:{minute:02d}!")
                break
            time.sleep(20)
    alarm_thread = threading.Thread(target=_a, daemon=True)
    alarm_thread.start()
    speak(f"Будильник на {hour:02d}:{minute:02d}.")
    return True

def stop_alarm():
    global alarm_thread; alarm_thread = None; return True

def break_reminder_on():
    global break_reminder_active
    if break_reminder_active:
        speak("Уже включено.")
        return True
    break_reminder_active = True
    def _r():
        while break_reminder_active:
            time.sleep(1800)
            if break_reminder_active:
                speak("Ты работаешь 30 минут. Перерыв!")
    threading.Thread(target=_r, daemon=True).start()
    return True

def break_reminder_off():
    global break_reminder_active
    break_reminder_active = False
    return True

def shutdown():
    os.system("shutdown /s /t 10")
    speak("Выключаю через 10 секунд.")
    return True

def restart():
    os.system("shutdown /r /t 10")
    speak("Перезагружаю через 10 секунд.")
    return True

def sleep_pc():
    os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
    return True

def cancel_shutdown():
    os.system("shutdown /a")
    return True

def dictate():
    speak("Говори — напечатаю")
    t = listen_fn()
    if not t:
        return False
    try:
        import pyautogui, pyperclip
        pyperclip.copy(t)
        pyautogui.hotkey("ctrl","v")
        return True
    except Exception:
        return False

def find_file():
    speak("Что ищем?")
    name = listen_fn()
    if not name:
        return False
    try:
        r = subprocess.run(
            ["powershell.exe", f"Get-ChildItem -Path $env:USERPROFILE -Recurse -Filter '*{name}*' -ErrorAction SilentlyContinue | Select-Object -First 3 FullName | Format-Table -HideTableHeaders"],
            capture_output=True, text=True, timeout=10)
        found = [l.strip() for l in r.stdout.strip().splitlines() if l.strip()]
        speak(f"Нашла: {found[0]}" if found else f"{name} не найден.")
        return True
    except Exception:
        return False

def open_folder(folder_name: str = ""):
    folders = {
        "документы":    os.path.expanduser("~/Documents"),
        "загрузки":     os.path.expanduser("~/Downloads"),
        "рабочий стол": os.path.expanduser("~/Desktop"),
        "музыка":       os.path.expanduser("~/Music"),
        "изображения":  os.path.expanduser("~/Pictures"),
        "видео":        os.path.expanduser("~/Videos"),
    }
    if not folder_name:
        speak("Какую папку?")
        folder_name = listen_fn() or ""
    path = folders.get(folder_name.lower(), folder_name)
    if os.path.exists(path):
        os.startfile(path)
        return True
    return False

def calculate(expression: str = ""):
    replacements = {
        "плюс":"+","минус":"-","умножить на":"*","умножить":"*",
        "разделить на":"/","разделить":"/","в степени":"**",
    }
    expr = expression.lower()
    for w, op in replacements.items():
        expr = expr.replace(w, op)
    safe = "".join(c for c in expr if c in "0123456789+-*/(). ").strip()
    if not safe:
        return False
    try:
        result = eval(safe)
        if isinstance(result, float) and result == int(result):
            result = int(result)
        speak(f"{expression} равно {result}")
        return True
    except Exception:
        return False

def remind_me(minutes: int, text: str = "Напоминание!"):
    def _r():
        time.sleep(minutes * 60)
        speak(f"Напоминаю: {text}")
    threading.Thread(target=_r, daemon=True).start()
    speak(f"Напомню через {minutes} мин.")
    return True

def break_code():
    speak("До встречи!")
    exit()


def execute_command(cmd: str, query: str = "") -> bool:
    if cmd.startswith("open_folder:"):
        return open_folder(cmd.split(":",1)[1])
    if cmd.startswith("open_app:"):
        return open_app(cmd.split(":",1)[1])
    if cmd == "mute":
        global is_muted
        is_muted = True
        stop_speech()
        return True
    if cmd == "unmute":
        is_muted = False
        play_cached("listening_again")
        return True
    if cmd == "set_timer":
        nums = re.findall(r'\d+', query)
        return set_timer(int(nums[0]) * 60 if nums else 60)
    if cmd == "set_alarm":
        nums = re.findall(r'\d+', query)
        if len(nums) >= 2:
            return set_alarm(int(nums[0]), int(nums[1]))
        elif len(nums) == 1:
            return set_alarm(int(nums[0]), 0)
        return False
    if cmd == "calculate":
        return calculate(query)
    if cmd == "remind_me":
        nums = re.findall(r'\d+', query)
        return remind_me(int(nums[0]) if nums else 5, query)

    dispatch = {
        "get_time":          get_time,
        "get_date":          get_date,
        "screenshot":        screenshot,
        "clipboard_read":    clipboard_read,
        "volume_up":         volume_up,
        "volume_down":       volume_down,
        "sound_off":         sound_off,
        "sound_on":          sound_on,
        "brightness_up":     brightness_up,
        "brightness_down":   brightness_down,
        "switch_window":     switch_window,
        "window_minimize":   window_minimize,
        "window_maximize":   window_maximize,
        "window_close":      window_close,
        "clipboard_copy":    clipboard_copy,
        "clipboard_paste":   clipboard_paste,
        "clipboard_write":   clipboard_write,
        "get_battery":       get_battery,
        "get_cpu":           get_cpu,
        "wifi_on":           wifi_on,
        "wifi_off":          wifi_off,
        "get_weather":       get_weather,
        "play_music":        play_music,
        "stop_music":        stop_music,
        "create_task":       create_task,
        "show_tasks":        show_tasks,
        "clear_tasks":       clear_tasks,
        "open_browser":      open_browser,
        "close_browser":     close_browser,
        "find_file":         find_file,
        "open_folder":       open_folder,
        "dictate":           dictate,
        "shutdown":          shutdown,
        "restart":           restart,
        "sleep":             sleep_pc,
        "cancel_shutdown":   cancel_shutdown,
        "stop_alarm":        stop_alarm,
        "break_reminder_on": break_reminder_on,
        "break_reminder_off":break_reminder_off,
        "break_code":        break_code,
    }
    fn = dispatch.get(cmd)
    return fn() if fn else False


# ─── MAIN ────────────────────────────────────────────────────────────────────

def main():
    global listen_fn, is_muted, last_activation

    os.makedirs("music", exist_ok=True)
    os.makedirs("sounds", exist_ok=True)
    if not os.path.exists("список дел.txt"):
        open("список дел.txt", "w", encoding="utf-8").close()

    # Генерируем кеш фраз
    _generate_cache()

    # ── Слушалка после wake word ──
    # Используем Google SR как fallback — WSR занят грамматикой команд
    def _listen(timeout=8) -> str | None:
        return listen_google(timeout)
    listen_fn = _listen

    # ── Проверка Porcupine ──
    if not PICOVOICE_KEY:
        print("  [!] Добавь PICOVOICE_KEY в .env")
        exit(1)
    if not os.path.exists(PPN_PATH):
        print(f"  [!] Файл не найден: {PPN_PATH}")
        exit(1)

    # ── WSR грамматика ──
    all_phrases = list(LOCAL_COMMANDS.keys())
    if WSR_AVAILABLE:
        try:
            wsr = WSRListener(all_phrases)
            print(f"  [wsr] Грамматика загружена — {len(all_phrases)} фраз")
            use_wsr = True
        except Exception as e:
            print(f"  [!] WSR недоступен: {e}")
            use_wsr = False
    else:
        print("  [!] pywin32 не установлен — pip install pywin32")
        use_wsr = False

    # ── Porcupine ──
    porcupine = pvporcupine.create(
        access_key=PICOVOICE_KEY,
        keyword_paths=[PPN_PATH],
        sensitivities=[0.5]
    )
    recorder = PvRecorder(device_index=-1, frame_length=porcupine.frame_length)
    recorder.start()

    play_cached("ready")
    print("  Говори 'Эй Лора'.\n")

    def _process(query: str):
        global is_muted, last_activation

        if is_speaking:
            stop_speech()
            time.sleep(0.05)

        query = query.lower().strip()
        print(f"  [cmd] {query}")

        # Стоп-триггеры
        if any(w in query for w in STOP_TRIGGERS):
            break_code()

        # Мут
        if any(w in query for w in MUTE_TRIGGERS):
            is_muted = True
            stop_speech()
            return

        # Размут
        if is_muted:
            if any(w in query for w in UNMUTE_TRIGGERS):
                is_muted = False
                play_cached("listening_again")
            return

        # Ищем команду — точное совпадение (WSR уже выдал нужную фразу)
        cmd = LOCAL_COMMANDS.get(query)

        if cmd:
            result = execute_command(cmd, query)
            if result:
                play_cached_random("confirm")
            else:
                play_cached_random("unclear")
        else:
            # WSR не должен выдавать неизвестные фразы, но на всякий случай
            play_cached_random("unclear")

    try:
        while True:
            pcm = recorder.read()
            if porcupine.process(pcm) >= 0:
                recorder.stop()
                play_wake()
                last_active = time.time()
                time.sleep(0.15)

                # Слушаем команду
                if use_wsr:
                    cmd_text = wsr.listen(timeout=6)
                else:
                    cmd_text = listen_google(timeout=6)

                recorder.start()

                if not cmd_text:
                    continue

                print(f"  [you] {cmd_text}")
                _process(cmd_text)
                last_active = time.time()

                # Окно активности — слушаем продолжение без wake word
                while (time.time() - last_active) < WINDOW_AFTER_AI:
                    if use_wsr:
                        followup = wsr.listen(timeout=2)
                    else:
                        followup = listen_google(timeout=2)
                    if followup:
                        last_active = time.time()
                        print(f"  [you] {followup}")
                        _process(followup)

    except KeyboardInterrupt:
        pass
    finally:
        recorder.stop()
        recorder.delete()
        porcupine.delete()
        if use_wsr:
            wsr.close()


if __name__ == "__main__":
    main()
