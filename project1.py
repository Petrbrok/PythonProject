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
from groq import Groq
from translate import Translator
import edge_tts
import pygame
from fuzzywuzzy import fuzz

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

import speech_recognition

load_dotenv()

ai_client = Groq(api_key=os.getenv("GROQ_API_KEY"))

EDGE_VOICE = os.getenv("EDGE_VOICE", "ru-RU-SvetlanaNeural")
WEATHER_API_KEY = os.getenv("WEATHER_API_KEY", "")

pygame.mixer.init()
sr_engine = speech_recognition.Recognizer()
sr_engine.pause_threshold = 0.6

# ─────────────────────────── СОСТОЯНИЕ ───────────────────────────

is_muted          = False
is_speaking       = False
always_listen     = False
last_activation   = 0.0
stop_speaking_event = threading.Event()

WINDOW_AFTER_COMMAND = 10
WINDOW_AFTER_AI      = 15

alarm_thread          = None
break_reminder_active = False

# ─────────────────────────── КОНСТАНТЫ ───────────────────────────

NAME_TRIGGERS = ("лора", "флора", "laura", "лёра", "лаура", "лор", "клара", "хлора", "лара", "пора", "лоор", "лёра")

WAKE_PHRASES  = ["Слушаю.", "Да.", "Здесь."]
CONFIRM_PHRASES = ["Есть!", "Выполняю.", "Сделано.", "Готово.", "Принято."]
ERROR_PHRASES   = ["Не получилось, попробуй снова.", "Ошибка, попробуй ещё раз.", "Что-то пошло не так."]

MUTE_TRIGGERS = (
    "замолчи", "молчи", "тихо", "пауза", "не слушай",
    "заткнись", "хватит", "достаточно", "подожди",
    "погоди", "перестань", "не говори", "тишина", "умолкни"
)

UNMUTE_TRIGGERS = (
    "размут", "включись", "продолжай",
    "проснись", "вернись", "активируйся"
)

STOP_TRIGGERS = (
    "завершить работу", "заверши работу", "выключись",
    "завершись", "закройся", "выход", "пока",
    "до свидания", "отключись", "выключи себя"
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

# ─────────────────────────── ЛОКАЛЬНЫЙ СЛОВАРЬ КОМАНД ───────────────────────────
# Формат: "фраза" -> "команда"
# Чем больше фраз — тем лучше распознавание без ИИ

LOCAL_COMMANDS = {
    # Время и дата
    "который час":          "get_time",
    "сколько времени":      "get_time",
    "какое время":          "get_time",
    "время":                "get_time",
    "какое сегодня число":  "get_date",
    "какая дата":           "get_date",
    "дата":                 "get_date",
    "день недели":          "get_date",

    # Громкость
    "увеличь громкость":    "volume_up",
    "громче":               "volume_up",
    "сделай громче":        "volume_up",
    "прибавь громкость":    "volume_up",
    "уменьши громкость":    "volume_down",
    "тише":                 "volume_down",
    "сделай тише":          "volume_down",
    "убавь громкость":      "volume_down",
    "выключи звук":         "sound_off",
    "без звука":            "sound_off",
    "заглуши":              "sound_off",
    "включи звук":          "sound_on",
    "верни звук":           "sound_on",

    # Яркость
    "увеличь яркость":      "brightness_up",
    "ярче":                 "brightness_up",
    "сделай ярче":          "brightness_up",
    "уменьши яркость":      "brightness_down",
    "темнее":               "brightness_down",
    "сделай темнее":        "brightness_down",

    # Окна
    "сверни окно":          "window_minimize",
    "сверни все":           "window_minimize",
    "свернуть окно":        "window_minimize",
    "сверни все окна":      "window_minimize",
    "убери все окна":       "window_minimize",
    "скрой все окна":       "window_minimize",
    "сверни всё":           "window_minimize",
    "разверни окно":        "window_maximize",
    "развернуть окно":      "window_maximize",
    "закрой окно":          "window_close",
    "закрыть окно":         "window_close",
    "следующее окно":       "switch_window",
    "переключи окно":       "switch_window",
    "смени окно":           "switch_window",
    "альт таб":             "switch_window",

    # Буфер
    "скопируй":             "clipboard_copy",
    "скопировать":          "clipboard_copy",
    "вставь":               "clipboard_paste",
    "вставить":             "clipboard_paste",
    "что в буфере":         "clipboard_read",
    "прочитай буфер":       "clipboard_read",

    # Скриншот
    "скриншот":             "screenshot",
    "сделай скриншот":      "screenshot",
    "снимок экрана":        "screenshot",

    # Система
    "заряд батареи":        "get_battery",
    "сколько заряда":       "get_battery",
    "батарея":              "get_battery",
    "загрузка процессора":  "get_cpu",
    "нагрузка":             "get_cpu",
    "включи вайфай":        "wifi_on",
    "включи wifi":          "wifi_on",
    "выключи вайфай":       "wifi_off",
    "выключи wifi":         "wifi_off",

    # Приложения — браузер
    "открой браузер":       "open_browser",
    "закрой браузер":       "close_browser",

    # Музыка
    "включи музыку":        "play_music",
    "играй музыку":         "play_music",
    "выключи музыку":       "stop_music",
    "стоп музыка":          "stop_music",

    # Задачи
    "добавь задачу":        "create_task",
    "новая задача":         "create_task",
    "список задач":         "show_tasks",
    "мои задачи":           "show_tasks",
    "очисти задачи":        "clear_tasks",
    "удали задачи":         "clear_tasks",

    # Перевод
    "переведи":             "translate",
    "переведи фразу":       "translate",

    # Выключение
    "выключи компьютер":    "shutdown",
    "перезагрузи":          "restart",
    "перезагрузка":         "restart",
    "сон":                  "sleep",
    "спящий режим":         "sleep",
    "отмени выключение":    "cancel_shutdown",

    # Папки
    "открой загрузки":      "open_folder:загрузки",
    "открой документы":     "open_folder:документы",
    "открой рабочий стол":  "open_folder:рабочий стол",
    "открой музыку":        "open_folder:музыка",
    "открой видео":         "open_folder:видео",
    "открой изображения":   "open_folder:изображения",

    # Режимы
    "отвечай всегда":       "always_listen_on",
    "слушай всегда":        "always_listen_on",
    "не только имя":        "always_listen_on",
    "без имени":            "always_listen_on",
    "только имя":           "always_listen_off",
    "слушай только имя":    "always_listen_off",

    # Напоминания о перерывах
    "включи перерывы":      "break_reminder_on",
    "напоминай о перерывах": "break_reminder_on",
    "выключи перерывы":     "break_reminder_off",
}

# Загружаем дополнительные команды из commands.yaml если файл существует
_yaml_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "commands.yaml")
if os.path.exists(_yaml_path):
    try:
        _yaml_cmds = yaml.safe_load(open(_yaml_path, encoding="utf-8"))
        if _yaml_cmds:
            for _cmd, _phrases in _yaml_cmds.items():
                if isinstance(_phrases, list):
                    for _phrase in _phrases:
                        LOCAL_COMMANDS[_phrase.lower()] = _cmd
        print(f"  [yaml] Загружено команд из commands.yaml")
    except Exception as e:
        print(f"  [!] Ошибка commands.yaml: {e}")

# ─────────────────────────── НЕЧЁТКОЕ РАСПОЗНАВАНИЕ (fuzzywuzzy) ───────────────────────────

def levenshtein(s1: str, s2: str) -> int:
    """Оставляем для совместимости с поиском имени."""
    if len(s1) < len(s2):
        return levenshtein(s2, s1)
    if len(s2) == 0:
        return len(s1)
    prev = list(range(len(s2) + 1))
    for i, c1 in enumerate(s1):
        curr = [i + 1]
        for j, c2 in enumerate(s2):
            curr.append(min(prev[j + 1] + 1, curr[j] + 1,
                            prev[j] + (0 if c1 == c2 else 1)))
        prev = curr
    return prev[-1]

def find_local_command(query: str):
    """Ищет команду через fuzzywuzzy — как у Джарвиса."""
    q = query.lower().strip()
    if not q:
        return None

    best_cmd    = None
    best_score  = 0

    for phrase, cmd in LOCAL_COMMANDS.items():
        # комбинированный скор: ratio + partial_ratio (как в jarvis)
        score = (fuzz.ratio(q, phrase) * 0.6 +
                 fuzz.partial_ratio(q, phrase) * 0.4)
        if score > best_score:
            best_score = score
            best_cmd   = cmd

    # Порог 70 как у Джарвиса
    if best_score >= 70:
        print(f"  [fuzzy] {best_score:.0f}% → {best_cmd}")
        return best_cmd

    return None

# ─────────────────────────── РЕЧЬ ───────────────────────────

def _synthesize(text: str):
    async def _run():
        tts = edge_tts.Communicate(text, EDGE_VOICE)
        with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as f:
            tmp = f.name
        await tts.save(tmp)
        return tmp
    try:
        return asyncio.run(_run())
    except Exception as e:
        print(f"  [!] edge-tts: {type(e).__name__}, переключаюсь на офлайн")
        return None

def _speak_offline(text: str):
    try:
        import pyttsx3
        engine = pyttsx3.init()
        for voice in engine.getProperty("voices"):
            if "ru" in voice.id.lower() or "russian" in voice.name.lower():
                engine.setProperty("voice", voice.id)
                break
        engine.setProperty("rate", 160)
        engine.say(text)
        engine.runAndWait()
        engine.stop()
    except Exception as e:
        print(f"  [!] Офлайн голос: {e}")

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
            tmp = _synthesize(text)
            if tmp and not stop_speaking_event.is_set():
                pygame.mixer.music.load(tmp)
                pygame.mixer.music.play()
                while pygame.mixer.music.get_busy():
                    if stop_speaking_event.is_set():
                        pygame.mixer.music.stop()
                        break
                    pygame.time.Clock().tick(10)
                pygame.mixer.music.unload()
            elif not stop_speaking_event.is_set():
                _speak_offline(text)
        except Exception as e:
            print(f"  [!] Речь: {e}")
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

def play_sound(name="confirm"):
    for ext in ("mp3", "wav"):
        path = os.path.join("sounds", f"{name}.{ext}")
        if os.path.exists(path):
            try:
                sound = pygame.mixer.Sound(path)
                sound.play()
            except Exception:
                pass
            return

# Кеш базовых фраз — генерируются один раз, воспроизводятся мгновенно
CACHED_PHRASES = {
    "ready":          "Готова к работе",
    "listening_again":"Снова слушаю",
    "no_internet":    "Нет интернета, попробуй базовые команды",
    "error_0":        "Не получилось, попробуй снова",
    "error_1":        "Что-то пошло не так",
    "error_2":        "Ошибка, попробуй ещё раз",
    "confirm_0":      "Есть",
    "confirm_1":      "Выполняю",
    "confirm_2":      "Сделано",
    "confirm_3":      "Готово",
    "confirm_4":      "Принято",
    "unclear_0":      "Не поняла",
    "unclear_1":      "Повтори пожалуйста",
    "unclear_2":      "Не расслышала",
}

_phrase_cache: dict[str, str] = {}  # key -> путь к mp3

def _generate_wake_sounds():
    """Генерирует wake-фразы и базовые фразы при первом запуске."""
    wake_dir = os.path.join("sounds", "wake")
    cache_dir = os.path.join("sounds", "cache")
    os.makedirs(wake_dir, exist_ok=True)
    os.makedirs(cache_dir, exist_ok=True)

    async def _gen(text, path):
        tts = edge_tts.Communicate(text, EDGE_VOICE)
        await tts.save(path)

    # Wake фразы
    for i, phrase in enumerate(WAKE_PHRASES):
        path = os.path.join(wake_dir, f"wake_{i}.mp3")
        if not os.path.exists(path):
            print(f"  [cache] Генерирую: {phrase}")
            try:
                asyncio.run(_gen(phrase, path))
            except Exception as e:
                print(f"  [!] {e}")

    # Базовые фразы
    for key, phrase in CACHED_PHRASES.items():
        path = os.path.join(cache_dir, f"{key}.mp3")
        if not os.path.exists(path):
            print(f"  [cache] Генерирую: {phrase}")
            try:
                asyncio.run(_gen(phrase, path))
            except Exception as e:
                print(f"  [!] {e}")
        if os.path.exists(path):
            _phrase_cache[key] = path

def _play_file(path: str):
    """Воспроизводит mp3 файл мгновенно."""
    try:
        pygame.mixer.music.load(path)
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy():
            pygame.time.Clock().tick(10)
        pygame.mixer.music.unload()
    except Exception:
        pass

def play_cached(key: str):
    """Воспроизводит кешированную фразу мгновенно."""
    if key in _phrase_cache:
        _play_file(_phrase_cache[key])
    else:
        # Fallback на текст
        text = CACHED_PHRASES.get(key, "")
        if text:
            speak(text)

def play_cached_random(prefix: str):
    """Воспроизводит случайную кешированную фразу с заданным префиксом."""
    keys = [k for k in _phrase_cache if k.startswith(prefix)]
    if keys:
        _play_file(_phrase_cache[random.choice(keys)])
    else:
        if prefix == "confirm":
            speak(random.choice(CONFIRM_PHRASES))
        elif prefix == "error":
            speak(random.choice(ERROR_PHRASES))

def play_wake():
    """Воспроизводит случайную wake-фразу мгновенно."""
    wake_dir = os.path.join("sounds", "wake")
    files = [f for f in os.listdir(wake_dir) if f.endswith(".mp3")] if os.path.exists(wake_dir) else []
    if files:
        _play_file(os.path.join(wake_dir, random.choice(files)))
    else:
        speak(random.choice(WAKE_PHRASES))

# ─────────────────────────── МИК ───────────────────────────

def _init_vosk():
    if not VOSK_AVAILABLE:
        return None
    model_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "model")
    if not os.path.exists(model_path):
        print("  [!] Папка model/ не найдена — Google SR")
        return None
    try:
        vosk.SetLogLevel(-1)
        model = vosk.Model(model_path)
        print("  [vosk] Модель загружена")
        return model
    except Exception as e:
        print(f"  [!] Vosk: {e}")
        return None

def listen_vosk(vosk_model, timeout=8):
    """Слушает через Vosk + VAD."""
    if not SD_AVAILABLE:
        return listen_google(timeout)

    rec = vosk.KaldiRecognizer(vosk_model, 16000)
    q   = queue.Queue()
    speech_detected = False
    silence_frames  = 0
    SILENCE_LIMIT   = 20  # ~0.5 сек тишины после речи

    def _callback(indata, frames, time_info, status):
        if not is_speaking:
            q.put(bytes(indata))

    try:
        with sd.RawInputStream(samplerate=16000, blocksize=4000,
                               dtype="int16", channels=1, callback=_callback):
            start = time.time()
            while time.time() - start < timeout:
                try:
                    data = q.get(timeout=0.1)
                except queue.Empty:
                    continue

                if is_speaking:
                    stop_speech()

                if rec.AcceptWaveform(data):
                    result = json.loads(rec.Result())
                    text = result.get("text", "").strip()
                    if text:
                        print(f"  [you] {text}")
                        return text
                else:
                    partial = json.loads(rec.PartialResult()).get("partial", "")
                    if partial:
                        speech_detected = True
                        silence_frames  = 0
                    elif speech_detected:
                        silence_frames += 1
                        if silence_frames >= SILENCE_LIMIT:
                            # Финализируем
                            result = json.loads(rec.FinalResult())
                            text = result.get("text", "").strip()
                            if text:
                                print(f"  [you] {text}")
                                return text
                            speech_detected = False
                            silence_frames  = 0
                            rec = vosk.KaldiRecognizer(vosk_model, 16000)
    except Exception as e:
        print(f"  [!] Vosk ошибка: {e}")
        return listen_google(timeout)
    return None

def listen_google(timeout=None):
    """Google SR как fallback."""
    while True:
        try:
            with speech_recognition.Microphone() as source:
                audio = sr_engine.listen(source, phrase_time_limit=8, timeout=timeout)
            if is_speaking:
                stop_speech()
            query = sr_engine.recognize_google(audio, language="ru-RU").lower().strip()
            print(f"  [you] {query}")
            return query
        except speech_recognition.WaitTimeoutError:
            return None
        except speech_recognition.UnknownValueError:
            pass
        except Exception:
            pass

# ─────────────────────────── ИИ ───────────────────────────

AI_SYSTEM_PROMPT = """Тебя зовут Лора — голосовой ассистент с дружелюбным характером. Общаешься как друг — просто, с лёгким юмором, без официоза. Отвечай на русском языке.

Длина ответа:
- Простой вопрос (факт, да/нет) — 1 предложение.
- Просят рассказать, объяснить, пошутить, написать творческое — отвечай столько сколько нужно.

КАТЕГОРИЧЕСКИ ЗАПРЕЩЕНО упоминать время, дату или время суток ("сейчас вечер", "сейчас утро" и т.д.) если тебя об этом явно не спросили. Нарушение этого правила недопустимо.
Когда называешь время — говори только "Сейчас [время]", без UTC и без "сейчас вечер/утро/день".

Если тебя спрашивают кто ты — отвечай что ты Лора, голосовой ассистент, без технических деталей.
Отвечай ТОЛЬКО текстом. Никакого JSON. Никаких технических меток."""

def ask_ai(query: str) -> str:
    now = datetime.now()
    try:
        r = ai_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            max_tokens=400,
            timeout=8,
            messages=[
                {"role": "system", "content": AI_SYSTEM_PROMPT},
                {"role": "user",   "content": f"[{now.strftime('%H:%M')}] {query}"},
            ],
        )
        return r.choices[0].message.content.strip()
    except Exception as e:
        return f"Ошибка связи: {e}"

def check_internet() -> bool:
    import urllib.request
    for url in ("https://www.bing.com", "https://www.cloudflare.com"):
        try:
            urllib.request.urlopen(url, timeout=3)
            return True
        except Exception:
            continue
    return False

# ─────────────────────────── КОМАНДЫ ───────────────────────────

def get_time():
    return f"Сейчас {datetime.now().strftime('%H:%M')}."

def get_date():
    months = ["января","февраля","марта","апреля","мая","июня",
              "июля","августа","сентября","октября","ноября","декабря"]
    now = datetime.now()
    return f"Сегодня {now.day} {months[now.month-1]} {now.year} года, {['понедельник','вторник','среда','четверг','пятница','суббота','воскресенье'][now.weekday()]}."

def screenshot():
    try:
        import pyautogui
        fname = f"screenshot_{int(time.time())}.png"
        pyautogui.screenshot(fname)
        return f"Скриншот сохранён: {fname}"
    except Exception as e:
        return None  # ошибка — скажем случайную фразу ошибки

def clipboard_read():
    try:
        import pyperclip
        text = pyperclip.paste()
        return f"В буфере: {text[:200]}" if text else "Буфер пуст"
    except Exception:
        return None

def _resolve_app(name: str) -> str:
    raw = APP_PATHS.get(name.lower(), name)
    return raw.replace("%USERNAME%", os.environ.get("USERNAME", ""))

def open_app(name: str):
    path = _resolve_app(name)
    try:
        subprocess.Popen(path)
        return True
    except Exception:
        try:
            subprocess.Popen(path, shell=True)
            return True
        except Exception:
            return False

def close_app(name: str):
    process_map = {
        "telegram":"telegram","телеграм":"telegram",
        "discord":"discord","дискорд":"discord",
        "spotify":"spotify","спотифай":"spotify",
        "chrome":"chrome","хром":"chrome",
        "notepad":"notepad","блокнот":"notepad",
        "word":"winword","excel":"excel","powerpoint":"powerpnt",
        "obs":"obs64","steam":"steam","vscode":"code","код":"code",
    }
    proc = process_map.get(name.lower(), name)
    try:
        subprocess.run(["powershell.exe", f"Stop-Process -Name {proc} -ErrorAction SilentlyContinue"], check=False)
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

def _get_volume_interface():
    try:
        from ctypes import POINTER, cast
        from comtypes import CLSCTX_ALL
        from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        return cast(interface, POINTER(IAudioEndpointVolume))
    except Exception:
        return None

def sound_off():
    vol = _get_volume_interface()
    if vol:
        vol.SetMute(1, None)
    else:
        os.system("powershell.exe (new-object -com wscript.shell).SendKeys([char]173)")
    return True

def sound_on():
    vol = _get_volume_interface()
    if vol:
        vol.SetMute(0, None)
    else:
        os.system("powershell.exe (new-object -com wscript.shell).SendKeys([char]173)")
    return True

def volume_up():
    vol = _get_volume_interface()
    if vol:
        current = vol.GetMasterVolumeLevelScalar()
        vol.SetMasterVolumeLevelScalar(min(1.0, current + 0.1), None)
    else:
        for _ in range(5):
            os.system("powershell.exe (new-object -com wscript.shell).SendKeys([char]175)")
    return True

def volume_down():
    vol = _get_volume_interface()
    if vol:
        current = vol.GetMasterVolumeLevelScalar()
        vol.SetMasterVolumeLevelScalar(max(0.0, current - 0.1), None)
    else:
        for _ in range(5):
            os.system("powershell.exe (new-object -com wscript.shell).SendKeys([char]174)")
    return True

def brightness_up():
    try:
        import screen_brightness_control as sbc
        cur = sbc.get_brightness()[0]
        sbc.set_brightness(min(100, cur + 20))
        return True
    except Exception:
        os.system("powershell (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,[math]::Min(100,((Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightness).CurrentBrightness)+20))")
        return True

def brightness_down():
    try:
        import screen_brightness_control as sbc
        cur = sbc.get_brightness()[0]
        sbc.set_brightness(max(0, cur - 20))
        return True
    except Exception:
        os.system("powershell (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,[math]::Max(0,((Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightness).CurrentBrightness)-20))")
        return True

def switch_window():
    import pyautogui
    pyautogui.hotkey("alt", "tab")
    return True

def window_minimize():
    import pyautogui
    pyautogui.hotkey("win", "d")
    return True

def window_maximize():
    import pyautogui
    pyautogui.hotkey("win", "up")
    return True

def window_close():
    import pyautogui
    pyautogui.hotkey("alt", "f4")
    return True

def clipboard_copy():
    import pyautogui
    pyautogui.hotkey("ctrl", "c")
    return True

def clipboard_paste():
    import pyautogui
    pyautogui.hotkey("ctrl", "v")
    return True

def clipboard_write():
    speak("Что записать?")
    text = listen_fn()
    if not text:
        return False
    try:
        import pyperclip
        pyperclip.copy(text)
        return True
    except Exception:
        return False

def get_battery():
    try:
        import psutil
        b = psutil.sensors_battery()
        if b is None:
            return "Батарея не найдена"
        status = "заряжается" if b.power_plugged else "на батарее"
        speak(f"Заряд {int(b.percent)}%, {status}.")
        return None
    except Exception:
        return None

def get_cpu():
    try:
        import psutil
        cpu = psutil.cpu_percent(interval=1)
        ram = psutil.virtual_memory().percent
        speak(f"Процессор {cpu}%, память {ram}%.")
        return None
    except Exception:
        return None

def wifi_off():
    os.system('netsh interface set interface "Wi-Fi" disabled')
    return True

def wifi_on():
    os.system('netsh interface set interface "Wi-Fi" enabled')
    return True

def get_weather():
    speak("Какой город?")
    city = listen_fn()
    if not city:
        return False
    if not WEATHER_API_KEY:
        speak("Добавь WEATHER_API_KEY в .env")
        return None
    try:
        import urllib.request, urllib.parse
        url = (f"https://api.openweathermap.org/data/2.5/weather"
               f"?q={urllib.parse.quote(city)}&appid={WEATHER_API_KEY}&units=metric&lang=ru")
        with urllib.request.urlopen(url, timeout=5) as r:
            data = json.loads(r.read().decode())
        desc  = data["weather"][0]["description"]
        temp  = round(data["main"]["temp"])
        feels = round(data["main"]["feels_like"])
        humid = data["main"]["humidity"]
        wind  = round(data["wind"]["speed"])
        speak(f"В {city}: {desc}, {temp}°, ощущается как {feels}°, влажность {humid}%, ветер {wind} м/с.")
        return None
    except Exception as e:
        speak(f"Не удалось получить погоду.")
        return None

def create_task():
    speak("Что добавить?")
    task = listen_fn()
    if not task:
        return False
    with open("список дел.txt", "a", encoding="utf-8") as f:
        f.write(f"✅ {task}\n")
    speak(f"Добавила: {task}")
    return None

def show_tasks():
    with open("список дел.txt", "r", encoding="utf-8") as f:
        tasks = f.read().strip()
    if not tasks:
        speak("Список пуст.")
    else:
        lines = tasks.splitlines()
        speak(f"Задач {len(lines)}: " + "; ".join(l.replace("✅ ","") for l in lines[:5]))
    return None

def clear_tasks():
    open("список дел.txt", "w", encoding="utf-8").close()
    return True

def play_music():
    files = [f for f in os.listdir("music") if f.endswith((".mp3",".wav",".flac"))]
    if not files:
        speak("В папке music нет файлов.")
        return None
    f = os.path.join("music", random.choice(files))
    os.startfile(f)
    speak(f"Включаю {os.path.splitext(os.path.basename(f))[0]}")
    return None

def stop_music():
    for p in ("wmplayer","vlc","spotify","groove","musicbee"):
        try:
            subprocess.run(["powershell.exe", f"Stop-Process -Name {p} -ErrorAction SilentlyContinue"], check=False)
        except Exception:
            pass
    return True

def translate_phrase():
    speak("Скажи фразу")
    text = listen_fn()
    if not text:
        return False
    try:
        t = Translator(from_lang="ru", to_lang="en")
        speak(f"По-английски: {t.translate(text)}")
        return None
    except Exception:
        return False

def set_timer(seconds: int):
    def _t():
        time.sleep(seconds)
        speak("Таймер сработал!")
    threading.Thread(target=_t, daemon=True).start()
    mins, secs = divmod(seconds, 60)
    speak(f"Таймер на {mins} мин {secs} сек." if mins else f"Таймер на {secs} сек.")
    return None

def set_alarm(hour: int, minute: int):
    global alarm_thread
    def _a():
        while True:
            now = datetime.now()
            if now.hour == hour and now.minute == minute:
                speak(f"Будильник! {hour:02d}:{minute:02d}!")
                break
            time.sleep(20)
    alarm_thread = threading.Thread(target=_a, daemon=True)
    alarm_thread.start()
    speak(f"Будильник на {hour:02d}:{minute:02d}.")
    return None

def stop_alarm():
    global alarm_thread
    alarm_thread = None
    return True

def break_reminder_on():
    global break_reminder_active
    if break_reminder_active:
        speak("Уже включено.")
        return None
    break_reminder_active = True
    def _r():
        while break_reminder_active:
            time.sleep(30 * 60)
            if break_reminder_active:
                speak("Ты работаешь 30 минут. Время сделать перерыв!")
    threading.Thread(target=_r, daemon=True).start()
    return True

def break_reminder_off():
    global break_reminder_active
    break_reminder_active = False
    return True

def shutdown():
    os.system("shutdown /s /t 10")
    speak("Выключаю через 10 секунд.")
    return None

def restart():
    os.system("shutdown /r /t 10")
    speak("Перезагружаю через 10 секунд.")
    return None

def sleep_pc():
    os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
    return True

def cancel_shutdown():
    os.system("shutdown /a")
    return True

def dictate():
    speak("Говори — напечатаю")
    text = listen_fn()
    if not text:
        return False
    try:
        import pyautogui, pyperclip
        pyperclip.copy(text)
        pyautogui.hotkey("ctrl", "v")
        return True
    except Exception:
        return False

def find_file():
    speak("Что ищем?")
    name = listen_fn()
    if not name:
        return False
    try:
        result = subprocess.run(
            ["powershell.exe", f"Get-ChildItem -Path $env:USERPROFILE -Recurse -Filter '*{name}*' -ErrorAction SilentlyContinue | Select-Object -First 3 FullName | Format-Table -HideTableHeaders"],
            capture_output=True, text=True, timeout=10
        )
        found = [l.strip() for l in result.stdout.strip().splitlines() if l.strip()]
        if found:
            speak(f"Нашла {len(found)}: {found[0]}")
        else:
            speak(f"Файл {name} не найден.")
        return None
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
        speak("Какую папку открыть?")
        folder_name = listen_fn() or ""
    path = folders.get(folder_name.lower(), folder_name)
    if os.path.exists(path):
        os.startfile(path)
        return True
    return False

def notify(text: str = "Уведомление от Лоры"):
    try:
        subprocess.Popen(["powershell.exe", "-Command",
            f"[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime] | Out-Null; "
            f"$t = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent(0); "
            f"$t.GetElementsByTagName('text')[0].AppendChild($t.CreateTextNode('{text}')) | Out-Null; "
            f"[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('Лора').Show([Windows.UI.Notifications.ToastNotification]::new($t))"
        ])
    except Exception:
        pass
    return True

def remind_me(minutes: int, text: str = "Напоминание!"):
    def _r():
        time.sleep(minutes * 60)
        speak(f"Напоминаю: {text}")
        notify(text)
    threading.Thread(target=_r, daemon=True).start()
    speak(f"Напомню через {minutes} мин.")
    return None

def calculate(expression: str = ""):
    replacements = {
        "плюс":"+","минус":"-","умножить на":"*","умножить":"*",
        "разделить на":"/","разделить":"/","в степени":"**",
        "процентов от":"*0.01*","процент от":"*0.01*","процента от":"*0.01*",
    }
    expr = expression.lower()
    for word, op in replacements.items():
        expr = expr.replace(word, op)
    safe = "".join(c for c in expr if c in "0123456789+-*/(). ").strip()
    if not safe:
        return False
    try:
        result = eval(safe)
        if isinstance(result, float) and result == int(result):
            result = int(result)
        speak(f"{expression} равно {result}")
        return None
    except Exception:
        return False

def always_listen_on():
    global always_listen
    always_listen = True
    return True

def always_listen_off():
    global always_listen
    always_listen = False
    return True

def break_code():
    speak("До встречи!")
    exit()

# ─────────────────────────── ОБРАБОТКА КОМАНД ───────────────────────────

# Глобальная функция listen (устанавливается в main)
listen_fn = None

def execute_command(cmd: str, query: str = "") -> bool | None:
    """
    Выполняет команду.
    Возвращает: True/False — системная (показать confirm/error), None — команда сама озвучила результат
    """
    # Команды с параметрами из строки запроса
    if cmd.startswith("open_folder:"):
        return open_folder(cmd.split(":", 1)[1])

    dispatch = {
        "get_time":          lambda: speak(get_time()) or None,
        "get_date":          lambda: speak(get_date()) or None,
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
        "translate":         translate_phrase,
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
        "always_listen_on":  always_listen_on,
        "always_listen_off": always_listen_off,
        "break_code":        break_code,
    }

    # open_app / close_app — извлекаем название из запроса
    if cmd == "open_app":
        # Ищем название приложения в запросе
        for name in APP_PATHS:
            if name in query:
                return open_app(name)
        speak("Какое приложение открыть?")
        app = listen_fn()
        return open_app(app) if app else False

    if cmd == "close_app":
        for name in APP_PATHS:
            if name in query:
                return close_app(name)
        speak("Какое приложение закрыть?")
        app = listen_fn()
        return close_app(app) if app else False

    # set_timer — извлекаем число из запроса
    if cmd == "set_timer":
        nums = re.findall(r'\d+', query)
        seconds = int(nums[0]) * 60 if nums else 60
        set_timer(seconds)
        return None

    # set_alarm — извлекаем время
    if cmd == "set_alarm":
        nums = re.findall(r'\d+', query)
        if len(nums) >= 2:
            set_alarm(int(nums[0]), int(nums[1]))
        elif len(nums) == 1:
            set_alarm(int(nums[0]), 0)
        else:
            speak("Скажи на какое время поставить будильник.")
        return None

    # calculate
    if cmd == "calculate":
        calculate(query)
        return None

    # remind_me
    if cmd == "remind_me":
        nums = re.findall(r'\d+', query)
        minutes = int(nums[0]) if nums else 5
        remind_me(minutes, query)
        return None

    fn = dispatch.get(cmd)
    if fn:
        return fn()
    return False


# ─────────────────────────── MAIN ───────────────────────────

def main():
    global is_muted, last_activation, listen_fn

    print("\n  ╔══════════════════════════════╗")
    print("  ║   АССИСТЕНТ ЛОРА ЗАПУЩЕН     ║")
    print("  ╚══════════════════════════════╝\n")

    # Инициализация папок
    for folder in ("music", "sounds"):
        os.makedirs(folder, exist_ok=True)
    if not os.path.exists("список дел.txt"):
        open("список дел.txt", "w", encoding="utf-8").close()

    # Vosk
    vosk_model = _init_vosk()

    if not vosk_model:
        print("  Калибровка микрофона...")
        with speech_recognition.Microphone() as source:
            sr_engine.adjust_for_ambient_noise(source, duration=1.5)

    # Устанавливаем глобальную функцию listen
    def _listen(timeout=8):
        if vosk_model:
            return listen_vosk(vosk_model, timeout)
        return listen_google(timeout)
    listen_fn = _listen

    # Загружаем образцы голоса если есть
    samples_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "voice_samples.json")
    if os.path.exists(samples_path):
        try:
            samples = json.load(open(samples_path, encoding="utf-8"))
            # Добавляем варианты имени
            if "лора" in samples:
                for v in samples["лора"]:
                    if v and v not in NAME_TRIGGERS:
                        NAME_TRIGGERS = NAME_TRIGGERS + (v,)
            # Добавляем варианты мута
            if "мут" in samples:
                for v in samples["мут"]:
                    if v and v not in MUTE_TRIGGERS:
                        MUTE_TRIGGERS = MUTE_TRIGGERS + (v,)
            # Добавляем варианты размута
            if "размут" in samples:
                for v in samples["размут"]:
                    if v and v not in UNMUTE_TRIGGERS:
                        UNMUTE_TRIGGERS = UNMUTE_TRIGGERS + (v,)
            print(f"  [samples] Загружены образцы голоса")
        except Exception as e:
            print(f"  [!] Ошибка загрузки образцов: {e}")

    # Генерируем wake-фразы
    print("  Генерирую wake-фразы...")
    _generate_wake_sounds()

    print("  Готово. Говорите.\n")
    play_cached("ready")

    while True:
        query = listen_fn()
        if not query:
            continue

        if is_speaking:
            stop_speech()
            time.sleep(0.05)

        # Мут — без имени
        if is_muted:
            if any(w in query for w in UNMUTE_TRIGGERS):
                is_muted = False
                play_cached("listening_again")
            continue

        if any(w in query for w in MUTE_TRIGGERS):
            is_muted = True
            stop_speech()
            print("  [muted]")
            continue

        # Стоп — контекстно
        if any(w in query for w in STOP_TRIGGERS):
            if is_speaking:
                stop_speech()
                continue
            else:
                break_code()

        # Стоп — работает без имени
        if any(w in query for w in STOP_TRIGGERS):
            break_code()

        # Проверяем активацию через Левенштейн
        def _has_name(q):
            words = q.lower().split()
            for word in words:
                for name in NAME_TRIGGERS:
                    max_l = max(len(word), len(name))
                    if max_l == 0:
                        continue
                    score = 1 - levenshtein(word, name) / max_l
                    if score >= 0.75:
                        return True, name
            return False, None

        has_name, matched_name = _has_name(query)
        in_window = (time.time() - last_activation) < WINDOW_AFTER_AI

        if not has_name and not always_listen and not in_window:
            continue

        # Убираем имя из запроса
        clean = query
        if has_name and matched_name:
            # Убираем найденное слово которое похоже на имя
            words = clean.split()
            filtered = []
            removed = False
            for word in words:
                max_l = max(len(word), len(matched_name))
                score = 1 - levenshtein(word, matched_name) / max_l if max_l else 0
                if score >= 0.75 and not removed:
                    removed = True
                    continue
                filtered.append(word)
            clean = " ".join(filtered).strip(",. ")

        # Wake — воспроизводим фразу и ждём команду
        if not clean:
            play_wake()
            last_activation = time.time()
            cmd_query = listen_fn(timeout=8)
            if not cmd_query:
                continue
            clean = cmd_query

        # Фильтр мусора — игнорируем слишком короткие фразы
        if len(clean) < 3:
            play_wake()
            last_activation = time.time()
            cmd_query = listen_fn(timeout=8)
            if not cmd_query:
                continue
            clean = cmd_query

        print(f"  [cmd] {clean}")

        # ── Локальный поиск команды (мгновенно, без ИИ) ──
        local_cmd = find_local_command(clean)

        if local_cmd:
            result = execute_command(local_cmd, clean)
            if result is True:
                play_cached_random("confirm")
            elif result is False:
                play_cached_random("error")
            # None — команда сама озвучила
            last_activation = time.time() + (WINDOW_AFTER_COMMAND - WINDOW_AFTER_AI)

        else:
            # Фильтр мусора
            if len(clean) < 4 or (len(clean.split()) < 2 and len(clean) < 6):
                play_cached_random("unclear")
                continue

            # ── Groq — только для вопросов и неизвестных фраз ──
            online = check_internet()
            if not online:
                play_cached("no_internet")
                continue
            _t = time.time()
            response = ask_ai(clean)
            print(f"  [⏱] {time.time()-_t:.2f} сек")
            speak(response)
            last_activation = time.time()


if __name__ == "__main__":
    main()
