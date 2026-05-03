import os
import json
import time
import random
import asyncio
import tempfile
import subprocess
import threading
import webbrowser
import re
import queue
import yaml
from datetime import datetime
from dotenv import load_dotenv
import edge_tts
import pygame
import sounddevice as sd
import vosk

import keyboard
load_dotenv()

EDGE_VOICE      = os.getenv("EDGE_VOICE", "ru-RU-SvetlanaNeural")
WEATHER_API_KEY = os.getenv("WEATHER_API_KEY", "")
MODEL_PATH      = os.path.join(os.path.dirname(os.path.abspath(__file__)), "model")

try:
    from groq import Groq
    _groq_client = Groq(api_key=os.getenv("GROQ_API_KEY")) if os.getenv("GROQ_API_KEY") else None
    AI_ENABLED = _groq_client is not None
except ImportError:
    _groq_client = None
    AI_ENABLED = False

pygame.mixer.init(frequency=48000)

is_muted              = False
is_speaking           = False
stop_speaking_event   = threading.Event()
alarm_thread          = None
break_reminder_active = False
WINDOW_AFTER_AI       = 12
_first_activation     = True
_vosk_listener        = None

WAKE_PHRASES  = ["Слушаю.", "Да.", "Здесь."]
WAKE_WORDS    = ["эй лора", "эй лора", "лора слушай"]   # фразы для Vosk wake
MUTE_TRIGGERS = (
    "замолчи", "молчи", "тихо", "пауза", "не слушай",
    "заткнись", "хватит", "подожди", "погоди", "тишина",
    "умолкни", "мут", "помолчи",
)
UNMUTE_TRIGGERS = (
    "размут", "включись", "продолжай", "проснись", "вернись",
    "активируйся", "слушай",
)
STOP_TRIGGERS = (
    "завершить работу", "заверши работу", "выключись", "завершись",
    "закройся", "выход", "пока", "до свидания", "отключись", "выключи себя",
    "заверши код", "заверши кода", "завершить код", "закрой себя", "выключи лору",
)

CACHED_PHRASES = {
    "ready":     "Готова к работе",
    "unclear_0": "Не поняла",
    "unclear_1": "Повтори пожалуйста",
    "unclear_2": "Не расслышала",
}
_phrase_cache: dict = {}
_resp_cache:   dict = {}

ALL_RESPONSES = {
    "volume_up_10":"Громкость 10%.", "volume_up_15":"Громкость 15%.",
    "volume_up_20":"Громкость 20%.", "volume_up_25":"Громкость 25%.",
    "volume_up_30":"Громкость 30%.", "volume_up_35":"Громкость 35%.",
    "volume_up_40":"Громкость 40%.", "volume_up_45":"Громкость 45%.",
    "volume_up_50":"Громкость 50%.", "volume_up_55":"Громкость 55%.",
    "volume_up_60":"Громкость 60%.", "volume_up_65":"Громкость 65%.",
    "volume_up_70":"Громкость 70%.", "volume_up_75":"Громкость 75%.",
    "volume_up_80":"Громкость 80%.", "volume_up_85":"Громкость 85%.",
    "volume_up_90":"Громкость 90%.", "volume_up_95":"Громкость 95%.",
    "volume_up_100":"Громкость 100%.",
    "volume_max":"Громкость максимальная.", "volume_min":"Громкость минимальная.",
    "sound_off":"Звук выключен.", "sound_on":"Звук включён.",
    "volume_up_fallback":"Громкость увеличена.", "volume_down_fallback":"Громкость уменьшена.",
    "brightness_up":"Яркость увеличена.", "brightness_down":"Яркость уменьшена.",
    "brightness_20":"Яркость 20%.", "brightness_40":"Яркость 40%.",
    "brightness_60":"Яркость 60%.", "brightness_80":"Яркость 80%.",
    "brightness_100":"Яркость 100%.",
    "wifi_on":"WiFi включён.", "wifi_off":"WiFi выключен.",
    "window_minimize":"Окна свёрнуты.", "window_maximize":"Окно развёрнуто.",
    "window_close":"Окно закрыто.", "switch_window":"Переключаю.",
    "clipboard_copy":"Скопировано.", "clipboard_paste":"Вставлено.",
    "clipboard_empty":"Буфер пуст.", "screenshot":"Скриншот сохранён.",
    "stop_music":"Музыка остановлена.",
    "tasks_empty":"Список пуст.", "tasks_cleared":"Список очищен.",
    "shutdown":"Выключаю через 10 секунд.", "restart":"Перезагружаю через 10 секунд.",
    "sleep":"Спящий режим.", "cancel_shutdown":"Выключение отменено.",
    "alarm_off":"Будильник отключён.",
    "break_on":"Напоминания о перерывах включены.",
    "break_off":"Напоминания выключены.",
    "break_already":"Напоминания уже включены.",
    "browser_closed":"Браузер закрыт.",
    "app_telegram":"Телеграм открыт.", "app_discord":"Дискорд открыт.",
    "app_spotify":"Спотифай открыт.", "app_chrome":"Хром открыт.",
    "app_notepad":"Блокнот открыт.", "app_calculator":"Калькулятор открыт.",
    "app_explorer":"Проводник открыт.", "app_steam":"Стим открыт.",
    "app_obs":"OBS открыт.", "app_vscode":"VS Code открыт.",
    "app_word":"Word открыт.", "app_excel":"Excel открыт.",
    "ping_0":"Да?", "ping_1":"Я здесь.", "ping_2":"Слушаю.",
    "unmute_0":"Снова слушаю.", "unmute_1":"Да?", "unmute_2":"Я здесь.",
    "mode_night":"Ночной режим. Яркость снижена, уведомления отключены.",
    "mode_presentation":"Режим презентации включён.",
    "folder_downloads":"Загрузки открыта.", "folder_documents":"Документы открыта.",
    "folder_desktop":"Рабочий стол открыт.", "folder_music":"Музыка открыта.",
    "folder_video":"Видео открыто.", "folder_images":"Изображения открыты.",
    "dictated":"Напечатала.", "farewell":"До встречи!",
    "wake_0":"Слушаю.", "wake_1":"Да.", "wake_2":"Здесь.",
}

APP_PATHS = {
    "telegram":  os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Roaming\Telegram Desktop\Telegram.exe"),
    "телеграм":  os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Roaming\Telegram Desktop\Telegram.exe"),
    "discord":   os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Local\Discord\Update.exe"),
    "дискорд":   os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Local\Discord\Update.exe"),
    "spotify":   os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Roaming\Spotify\Spotify.exe"),
    "спотифай":  os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Roaming\Spotify\Spotify.exe"),
    "chrome":    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    "хром":      r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    "notepad":   "notepad.exe",    "блокнот":    "notepad.exe",
    "calculator":"calc.exe",       "калькулятор":"calc.exe",
    "paint":     "mspaint.exe",
    "проводник": "explorer.exe",   "explorer":   "explorer.exe",
    "word":      r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    "excel":     r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
    "powerpoint":r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
    "obs":       r"C:\Program Files\obs-studio\bin\64bit\obs64.exe",
    "steam":     r"C:\Program Files (x86)\Steam\steam.exe",
    "vscode":    os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Local\Programs\Microsoft VS Code\Code.exe"),
    "код":       os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Local\Programs\Microsoft VS Code\Code.exe"),
}
APP_NAMES_RU = {
    "телеграм":"Телеграм","telegram":"Телеграм",
    "дискорд":"Дискорд",  "discord":"Дискорд",
    "спотифай":"Спотифай","spotify":"Спотифай",
    "хром":"Хром",        "chrome":"Хром",
    "блокнот":"Блокнот",  "notepad":"Блокнот",
    "калькулятор":"Калькулятор","calculator":"Калькулятор",
    "проводник":"Проводник","explorer":"Проводник",
    "стим":"Стим",        "steam":"Стим",
    "obs":"OBS","код":"VS Code","vscode":"VS Code",
    "ворд":"Word","word":"Word","эксель":"Excel","excel":"Excel",
    "поверпоинт":"PowerPoint","powerpoint":"PowerPoint",
}

LOCAL_COMMANDS = {
    "который час":"get_time","сколько времени":"get_time",
    "какое время":"get_time","время":"get_time",
    "какое сегодня число":"get_date","какая дата":"get_date",
    "день недели":"get_date","дата":"get_date",
    "увеличь громкость":"volume_up","громче":"volume_up",
    "сделай громче":"volume_up","прибавь громкость":"volume_up",
    "уменьши громкость":"volume_down","тише":"volume_down",
    "сделай тише":"volume_down","убавь громкость":"volume_down",
    "выключи звук":"sound_off","без звука":"sound_off",
    "включи звук":"sound_on","верни звук":"sound_on",
    "максимальная громкость":"volume_max",
    "поставь максимальную громкость":"volume_max",
    "минимальная громкость":"volume_min",
    "увеличь яркость":"brightness_up","ярче":"brightness_up","сделай ярче":"brightness_up",
    "уменьши яркость":"brightness_down","темнее":"brightness_down","сделай темнее":"brightness_down",
    "сверни окно":"window_minimize","сверни все":"window_minimize",
    "сверни всё":"window_minimize","убери все окна":"window_minimize",
    "разверни окно":"window_maximize","закрой окно":"window_close",
    "переключи окно":"switch_window","альт таб":"switch_window",
    "скопируй":"clipboard_copy","вставь":"clipboard_paste","что в буфере":"clipboard_read",
    "скриншот":"screenshot","сделай скриншот":"screenshot","снимок экрана":"screenshot",
    "заряд батареи":"get_battery","сколько заряда":"get_battery","батарея":"get_battery",
    "загрузка процессора":"get_cpu","нагрузка системы":"get_cpu",
    "включи вайфай":"wifi_toggle_on","выключи вайфай":"wifi_toggle_off","вайфай":"wifi_toggle",
    "открой браузер":"open_browser","закрой браузер":"close_browser",
    "открой телеграм":"open_app:телеграм","запусти телеграм":"open_app:телеграм",
    "открой дискорд":"open_app:дискорд","запусти дискорд":"open_app:дискорд",
    "открой спотифай":"open_app:спотифай","запусти спотифай":"open_app:спотифай",
    "открой хром":"open_app:хром","запусти хром":"open_app:хром",
    "открой блокнот":"open_app:блокнот","открой калькулятор":"open_app:калькулятор",
    "открой проводник":"open_app:проводник","открой стим":"open_app:steam",
    "открой obs":"open_app:obs","открой код":"open_app:код",
    "открой ворд":"open_app:word","открой эксель":"open_app:excel",
    "включи музыку":"play_music","играй музыку":"play_music",
    "выключи музыку":"stop_music","стоп музыка":"stop_music",
    "добавь задачу":"create_task","новая задача":"create_task",
    "список задач":"show_tasks","мои задачи":"show_tasks","очисти задачи":"clear_tasks",
    "открой загрузки":"open_folder:загрузки",
    "открой документы":"open_folder:документы",
    "открой рабочий стол":"open_folder:рабочий стол",
    "открой музыку":"open_folder:музыка",
    "открой видео":"open_folder:видео",
    "открой изображения":"open_folder:изображения",
    "таймер":"set_timer","поставь таймер":"set_timer",
    "поставь будильник":"set_alarm","будильник":"stop_alarm","отключи будильник":"stop_alarm",
    "выключи компьютер":"shutdown","перезагрузи":"restart",
    "перезагрузка":"restart","спящий режим":"sleep","отмени выключение":"cancel_shutdown",
    "включи перерывы":"break_reminder_on","выключи перерывы":"break_reminder_off",
    "ночной режим":"mode_night","включи ночной режим":"mode_night",
    "режим презентации":"mode_presentation","презентация":"mode_presentation",
    "утренний режим":"mode_morning","доброе утро":"mode_morning",
    "список команд":"show_commands","покажи команды":"show_commands",
    "что ты умеешь":"show_commands","помощь":"show_commands",
    "замолчи":"mute","молчи":"mute","мут":"mute",
    "лора":"ping",
    "размут":"unmute","включись":"unmute","слушай":"unmute","продолжай":"unmute",
    "закрой телеграм":"close_app:телеграм","закрой дискорд":"close_app:дискорд",
    "закрой хром":"close_app:хром","закрой блокнот":"close_app:блокнот",
    "закрой ворд":"close_app:word","закрой эксель":"close_app:excel",
    "закрой стим":"close_app:steam",
    "погода":"get_weather","какая погода":"get_weather",
    "посчитай":"calculate","напомни":"remind_me",
}

_yaml_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "commands.yaml")
if os.path.exists(_yaml_path):
    try:
        _yaml_cmds = yaml.safe_load(open(_yaml_path, encoding="utf-8"))
        if _yaml_cmds:
            for _cmd, _phrases in _yaml_cmds.items():
                if isinstance(_phrases, list):
                    for _p in _phrases:
                        LOCAL_COMMANDS[_p.lower()] = _cmd
        print("  [yaml] Команды загружены")
    except Exception as e:
        print(f"  [!] commands.yaml: {e}")


# ─── VOSK ───────────────────────────────────────────────────────────────────

class VoskListener:
    SAMPLE_RATE = 16000
    BLOCK_SIZE  = 4000

    def __init__(self, phrases: list):
        if not os.path.exists(MODEL_PATH):
            raise RuntimeError(f"Модель Vosk не найдена: {MODEL_PATH}")
        vosk.SetLogLevel(-1)
        self._model    = vosk.Model(MODEL_PATH)
        # Без грамматики — слышит свободную речь, rapidfuzz сопоставляет с командами
        self._rec      = vosk.KaldiRecognizer(self._model, self.SAMPLE_RATE)
        self._rec_free = vosk.KaldiRecognizer(self._model, self.SAMPLE_RATE)
        # Wake word recognizer — только wake фразы
        wake_grammar   = json.dumps(WAKE_WORDS + ["[unk]"], ensure_ascii=False)
        self._rec_wake = vosk.KaldiRecognizer(self._model, self.SAMPLE_RATE, wake_grammar)
        print(f"  [vosk] Готов, фраз: {len(phrases)}")

    def listen(self, timeout: float = 8) -> str | None:
        self._rec.Reset()
        return self._listen_with(self._rec, timeout)

    def listen_free(self, timeout: float = 8) -> str | None:
        self._rec_free.Reset()
        return self._listen_with(self._rec_free, timeout)

    def _listen_with(self, rec, timeout: float) -> str | None:
        q  = queue.Queue()
        t0 = time.time()

        def _cb(indata, frames, t, status):
            q.put(bytes(indata))

        with sd.RawInputStream(samplerate=self.SAMPLE_RATE, blocksize=self.BLOCK_SIZE,
                               dtype="int16", channels=1, callback=_cb):
            while time.time() - t0 < timeout:
                try:
                    data = q.get(timeout=0.05)
                except queue.Empty:
                    continue
                if rec.AcceptWaveform(data):
                    text = json.loads(rec.Result()).get("text", "").strip()
                    if text and text != "[unk]":
                        return text

        text = json.loads(rec.FinalResult()).get("text", "").strip()
        return text if text and text != "[unk]" else None


class WakeWordListener:
    """
    Слушает wake word в фоновом потоке через Vosk с узкой грамматикой.
    Намного надёжнее DTW — нейросеть знает русский язык.
    """
    SAMPLE_RATE = 16000
    BLOCK_SIZE  = 2000   # ~125мс — быстрый отклик

    def __init__(self, model, callback):
        self._model    = model
        self._callback = callback
        self._running  = False
        self._paused   = False
        self._thread   = None
        grammar        = json.dumps(WAKE_WORDS + ["[unk]"], ensure_ascii=False)
        self._grammar  = grammar

    def start(self):
        self._running = True
        self._thread  = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()
        print("  [wake] Слушаю wake word...")

    def stop(self):
        self._running = False

    def pause(self):
        self._paused = True

    def resume(self):
        self._paused = False

    def _loop(self):
        rec = vosk.KaldiRecognizer(self._model, self.SAMPLE_RATE, self._grammar)
        q   = queue.Queue()

        def _cb(indata, frames, t, status):
            if not self._paused:
                q.put(bytes(indata))

        with sd.RawInputStream(samplerate=self.SAMPLE_RATE, blocksize=self.BLOCK_SIZE,
                               dtype="int16", channels=1, callback=_cb):
            while self._running:
                try:
                    data = q.get(timeout=0.1)
                except queue.Empty:
                    continue

                if rec.AcceptWaveform(data):
                    text = json.loads(rec.Result()).get("text", "").strip()
                    if text and text != "[unk]":
                        rec.Reset()
                        self._callback()
                else:
                    partial = json.loads(rec.PartialResult()).get("partial", "")
                    if any(w in partial for w in ["эй", "лора"]):
                        pass  # идёт распознавание


# ─── TTS — только edge-tts кеш ───────────────────────────────────────────────

def _generate_cache():
    os.makedirs(os.path.join("sounds","cache"), exist_ok=True)
    os.makedirs(os.path.join("sounds","wake"),  exist_ok=True)
    os.makedirs(os.path.join("sounds","resp"),  exist_ok=True)

    async def _gen(text, path):
        await edge_tts.Communicate(text, EDGE_VOICE, rate="+10%").save(path)

    for i, phrase in enumerate(WAKE_PHRASES):
        path = os.path.join("sounds","wake", f"wake_{i}.mp3")
        if not os.path.exists(path):
            try: asyncio.run(_gen(phrase, path))
            except Exception: pass

    for key, phrase in CACHED_PHRASES.items():
        path = os.path.join("sounds","cache", f"{key}.mp3")
        if not os.path.exists(path):
            try: asyncio.run(_gen(phrase, path))
            except Exception: pass
        if os.path.exists(path):
            _phrase_cache[key] = path

    total, new_count = len(ALL_RESPONSES), 0
    for key, phrase in ALL_RESPONSES.items():
        path = os.path.join("sounds","resp", f"{key}.mp3")
        if not os.path.exists(path):
            try: asyncio.run(_gen(phrase, path)); new_count += 1
            except Exception: pass
        if os.path.exists(path):
            _resp_cache[key] = path

    if new_count:
        print(f"  [cache] Сгенерировано {new_count} новых фраз")
    print(f"  [cache] Загружено {len(_resp_cache)}/{total} ответов")


def _play_file(path: str):
    try:
        pygame.mixer.music.load(path)
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy():
            pygame.time.Clock().tick(10)
        pygame.mixer.music.unload()
    except Exception:
        pass


def speak_resp(key: str) -> bool:
    if key in _resp_cache:
        _log_lora(ALL_RESPONSES[key])
        _play_file(_resp_cache[key])
        return True
    return False


def play_unclear():
    keys = [k for k in _phrase_cache if k.startswith("unclear")]
    if keys:
        chosen = random.choice(keys)
        _log_lora(CACHED_PHRASES.get(chosen, chosen))
        _play_file(_phrase_cache[chosen])
    else:
        speak(random.choice(["Не поняла", "Повтори пожалуйста", "Не расслышала"]))


def play_wake():
    keys = [k for k in _resp_cache if k.startswith("wake_")]
    if keys:
        chosen = random.choice(keys)
        _log_lora(ALL_RESPONSES[chosen])
        _play_file(_resp_cache[chosen])
    else:
        files = [f for f in os.listdir(os.path.join("sounds","wake")) if f.endswith(".mp3")] \
                if os.path.exists(os.path.join("sounds","wake")) else []
        if files:
            _play_file(os.path.join("sounds","wake", random.choice(files)))


def speak(text: str):
    """Динамические фразы через edge-tts (время, погода, ИИ)."""
    global is_speaking
    if not text: return
    _log_lora(text)
    is_speaking = True
    stop_speaking_event.clear()
    tmp = None
    try:
        async def _synth():
            tts = edge_tts.Communicate(text, EDGE_VOICE, rate="+10%")
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
    except Exception:
        pass
    finally:
        is_speaking = False
        if tmp:
            try: os.unlink(tmp)
            except Exception: pass


def stop_speech():
    global is_speaking
    stop_speaking_event.set()
    try: pygame.mixer.music.stop()
    except Exception: pass
    is_speaking = False


# ─── КОНСОЛЬ ────────────────────────────────────────────────────────────────

def _log_you(text: str): print(f"  you   {text}")
def _log_lora(text: str, ms: int = None):
    print(f"  lora  {text}" + (f"  [{ms}мс]" if ms else ""))


# ─── ИИ ─────────────────────────────────────────────────────────────────────

AI_SYSTEM_PROMPT = (
    "Тебя зовут Лора — голосовой ассистент. Отвечай по-русски, кратко и дружелюбно. "
    "Без сленга и грубостей. Простой вопрос — 1 предложение. Только текст, без эмодзи."
)

def ask_ai(query: str) -> str:
    if not AI_ENABLED or not _groq_client: return ""
    try:
        r = _groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile", max_tokens=150, timeout=8,
            messages=[
                {"role":"system","content":AI_SYSTEM_PROMPT},
                {"role":"user","content":f"[{datetime.now().strftime('%H:%M')}] {query}"},
            ],
        )
        return r.choices[0].message.content.strip()
    except Exception as e:
        return f"Ошибка: {e}"


# ─── СЛОТ-ЭКСТРАКТОР ────────────────────────────────────────────────────────

class SlotExtractor:
    @staticmethod
    def time_to_seconds(text: str) -> int | None:
        text = text.lower()
        if any(x in text for x in ["полчаса","пол часа"]): return 1800
        total = 0
        for pat, mul in [(r'(\d+)\s*час',3600),(r'(\d+)\s*мин',60),(r'(\d+)\s*сек',1)]:
            for m in re.findall(pat, text): total += int(m) * mul
        if total == 0:
            nums = re.findall(r'\d+', text)
            if nums: total = int(nums[0]) * 60
        return total if total > 0 else None


# ─── ГРОМКОСТЬ ───────────────────────────────────────────────────────────────

listen_fn = None

def _vol():
    try:
        from ctypes import POINTER, cast
        from comtypes import CLSCTX_ALL
        from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
        i = AudioUtilities.GetSpeakers().Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        return cast(i, POINTER(IAudioEndpointVolume))
    except Exception:
        return None

def _vol_init():
    v = _vol()
    if v:
        print(f"  [pycaw] Громкость: {int(round(v.GetMasterVolumeLevelScalar()*100))}%")
    else:
        print("  [!] pycaw недоступен — громкость через клавиши")

def _wifi_status() -> bool:
    try:
        r = subprocess.run(["netsh","interface","show","interface","name=Wi-Fi"],
                           capture_output=True, text=True, timeout=3)
        return "connect" in r.stdout.lower()
    except Exception: return False


# ─── КОМАНДЫ ────────────────────────────────────────────────────────────────

def get_time()  -> str: return f"Сейчас {datetime.now().strftime('%H:%M')}."
def get_date()  -> str:
    mo = ["января","февраля","марта","апреля","мая","июня",
          "июля","августа","сентября","октября","ноября","декабря"]
    n  = datetime.now()
    wd = ["понедельник","вторник","среда","четверг","пятница","суббота","воскресенье"][n.weekday()]
    return f"Сегодня {n.day} {mo[n.month-1]} {n.year}, {wd}."

def volume_up() -> str:
    v = _vol()
    if v:
        new = min(1.0, v.GetMasterVolumeLevelScalar() + 0.15)
        v.SetMasterVolumeLevelScalar(new, None)
        pct5 = round(round(new*100) / 5) * 5
        speak_resp(f"volume_up_{pct5}") or speak(f"Громкость {int(round(new*100))}%.")
    else:
        try:
            import pyautogui
            for _ in range(3): pyautogui.press("volumeup")
        except Exception: pass
        speak_resp("volume_up_fallback")
    return ""

def volume_down() -> str:
    v = _vol()
    if v:
        new = max(0.0, v.GetMasterVolumeLevelScalar() - 0.15)
        v.SetMasterVolumeLevelScalar(new, None)
        pct5 = round(round(new*100) / 5) * 5
        speak_resp(f"volume_up_{pct5}") or speak(f"Громкость {int(round(new*100))}%.")
    else:
        try:
            import pyautogui
            for _ in range(3): pyautogui.press("volumedown")
        except Exception: pass
        speak_resp("volume_down_fallback")
    return ""

def volume_max() -> str:
    v = _vol()
    if v: v.SetMasterVolumeLevelScalar(1.0, None)
    speak_resp("volume_max"); return ""

def volume_min() -> str:
    v = _vol()
    if v: v.SetMasterVolumeLevelScalar(0.0, None)
    speak_resp("volume_min"); return ""

def sound_off() -> str:
    v = _vol()
    if v: v.SetMute(1, None)
    else: os.system("powershell.exe (new-object -com wscript.shell).SendKeys([char]173)")
    speak_resp("sound_off"); return ""

def sound_on() -> str:
    v = _vol()
    if v: v.SetMute(0, None)
    else: os.system("powershell.exe (new-object -com wscript.shell).SendKeys([char]173)")
    speak_resp("sound_on"); return ""

def brightness_up() -> str:
    try:
        import screen_brightness_control as sbc
        new = min(100, sbc.get_brightness()[0] + 20)
        sbc.set_brightness(new)
        speak_resp(f"brightness_{round(new/20)*20}") or speak(f"Яркость {new}%.")
    except Exception:
        speak_resp("brightness_up")
    return ""

def brightness_down() -> str:
    try:
        import screen_brightness_control as sbc
        new = max(0, sbc.get_brightness()[0] - 20)
        sbc.set_brightness(new)
        speak_resp(f"brightness_{round(new/20)*20}") or speak(f"Яркость {new}%.")
    except Exception:
        speak_resp("brightness_down")
    return ""

def wifi_toggle_on()  -> str: os.system('netsh interface set interface "Wi-Fi" enabled');  speak_resp("wifi_on");  return ""
def wifi_toggle_off() -> str: os.system('netsh interface set interface "Wi-Fi" disabled'); speak_resp("wifi_off"); return ""
def wifi_toggle()     -> str:
    if _wifi_status(): return wifi_toggle_off()
    return wifi_toggle_on()

def screenshot() -> str:
    try:
        import pyautogui; pyautogui.screenshot(f"screenshot_{int(time.time())}.png")
        speak_resp("screenshot")
    except Exception: pass
    return ""

def clipboard_copy()  -> str:
    try: import pyautogui; pyautogui.hotkey("ctrl","c")
    except Exception: pass
    speak_resp("clipboard_copy"); return ""

def clipboard_paste() -> str:
    try: import pyautogui; pyautogui.hotkey("ctrl","v")
    except Exception: pass
    speak_resp("clipboard_paste"); return ""

def clipboard_read()  -> str:
    try:
        import pyperclip; t = pyperclip.paste()
        return f"В буфере: {t[:200]}" if t else "Буфер пуст."
    except Exception: return None

def switch_window()   -> str:
    try: import pyautogui; pyautogui.hotkey("alt","tab")
    except Exception: pass
    speak_resp("switch_window"); return ""

def window_minimize() -> str:
    try: import pyautogui; pyautogui.hotkey("win","d")
    except Exception: pass
    speak_resp("window_minimize"); return ""

def window_maximize() -> str:
    try: import pyautogui; pyautogui.hotkey("win","up")
    except Exception: pass
    speak_resp("window_maximize"); return ""

def window_close()    -> str:
    try: import pyautogui; pyautogui.hotkey("alt","f4")
    except Exception: pass
    speak_resp("window_close"); return ""

def get_battery() -> str:
    try:
        import psutil; b = psutil.sensors_battery()
        if b is None: return "Батарея не найдена."
        return f"Заряд {int(b.percent)}%, {'заряжается' if b.power_plugged else 'на батарее'}."
    except Exception: return None

def get_cpu() -> str:
    try:
        import psutil
        return f"Процессор {psutil.cpu_percent(interval=1)}%, память {psutil.virtual_memory().percent}%."
    except Exception: return None

def open_app(name: str) -> str:
    path    = APP_PATHS.get(name.lower(), name)
    ru_name = APP_NAMES_RU.get(name.lower(), name.capitalize())
    app_key_map = {
        "телеграм":"app_telegram","telegram":"app_telegram",
        "дискорд":"app_discord","discord":"app_discord",
        "спотифай":"app_spotify","spotify":"app_spotify",
        "хром":"app_chrome","chrome":"app_chrome",
        "блокнот":"app_notepad","notepad":"app_notepad",
        "калькулятор":"app_calculator","calculator":"app_calculator",
        "проводник":"app_explorer","explorer":"app_explorer",
        "стим":"app_steam","steam":"app_steam",
        "obs":"app_obs","код":"app_vscode","vscode":"app_vscode",
        "word":"app_word","ворд":"app_word","excel":"app_excel","эксель":"app_excel",
    }
    try:
        subprocess.Popen(path)
    except Exception:
        try: subprocess.Popen(path, shell=True)
        except Exception: return None
    key = app_key_map.get(name.lower())
    if not key or not speak_resp(key):
        speak(f"{ru_name} открыт.")
    return ""

def close_app(name: str) -> str:
    pm = {"telegram":"telegram","телеграм":"telegram","discord":"discord","дискорд":"discord",
          "spotify":"spotify","спотифай":"spotify","chrome":"chrome","хром":"chrome",
          "notepad":"notepad","блокнот":"notepad","word":"winword","excel":"excel",
          "obs":"obs64","steam":"steam","vscode":"code","код":"code"}
    ru_name = APP_NAMES_RU.get(name.lower(), name.capitalize())
    try:
        subprocess.run(["powershell.exe",
            f"Stop-Process -Name {pm.get(name.lower(),name)} -ErrorAction SilentlyContinue"],check=False)
        return f"{ru_name} закрыт."
    except Exception: return None

def open_browser() -> str:
    speak("Какой сайт?"); site = listen_fn()
    if not site: return None
    url = (f"https://{site}" if "." in site and " " not in site
           else f"https://www.google.com/search?q={site.replace(' ','+')}")
    webbrowser.open(url); return f"Открываю {site}."

def close_browser() -> str:
    for b in ("chrome","firefox","msedge","opera","brave"):
        try: subprocess.run(["powershell.exe",f"Stop-Process -Name {b} -ErrorAction SilentlyContinue"],check=False)
        except Exception: pass
    speak_resp("browser_closed"); return ""

def open_folder(folder_name: str = "") -> str:
    folders = {
        "документы":    os.path.expanduser("~/Documents"),
        "загрузки":     os.path.expanduser("~/Downloads"),
        "рабочий стол": os.path.expanduser("~/Desktop"),
        "музыка":       os.path.expanduser("~/Music"),
        "изображения":  os.path.expanduser("~/Pictures"),
        "видео":        os.path.expanduser("~/Videos"),
    }
    if not folder_name:
        speak("Какую папку?"); folder_name = listen_fn() or ""
    path = folders.get(folder_name.lower(), folder_name)
    if os.path.exists(path):
        os.startfile(path)
        key_map = {"загрузки":"folder_downloads","документы":"folder_documents",
                   "рабочий стол":"folder_desktop","музыка":"folder_music",
                   "видео":"folder_video","изображения":"folder_images"}
        speak_resp(key_map.get(folder_name.lower(), "")) or speak(f"{folder_name.capitalize()} открыта.")
        return ""
    return None

def get_weather() -> str:
    speak("Какой город?"); city = listen_fn()
    if not city: return None
    if not WEATHER_API_KEY: return "Добавь WEATHER_API_KEY в .env."
    try:
        import urllib.request, urllib.parse
        url = (f"https://api.openweathermap.org/data/2.5/weather"
               f"?q={urllib.parse.quote(city)}&appid={WEATHER_API_KEY}&units=metric&lang=ru")
        with urllib.request.urlopen(url, timeout=5) as r:
            d = json.loads(r.read().decode())
        desc  = d['weather'][0]['description']
        temp  = round(d['main']['temp'])
        feels = round(d['main']['feels_like'])
        return f"{city.capitalize()}: {desc}, {temp}°, ощущается {feels}°."
    except Exception: return "Не удалось получить погоду."

def create_task() -> str:
    speak("Что добавить?"); t = listen_fn()
    if not t: return None
    with open("список дел.txt","a",encoding="utf-8") as f: f.write(f"✅ {t}\n")
    return f"Добавила: {t}."

def show_tasks() -> str:
    try:
        with open("список дел.txt",encoding="utf-8") as f: tasks = f.read().strip()
        if not tasks: return "Список пуст."
        lines = tasks.splitlines()
        return f"Задач {len(lines)}: " + "; ".join(l.replace("✅ ","") for l in lines[:5]) + "."
    except Exception: return None

def clear_tasks() -> str:
    open("список дел.txt","w",encoding="utf-8").close()
    speak_resp("tasks_cleared"); return ""

def play_music() -> str:
    try:
        files = [f for f in os.listdir("music") if f.endswith((".mp3",".wav",".flac"))]
        if not files: return "В папке music нет файлов."
        f = os.path.join("music", random.choice(files))
        os.startfile(f)
        return f"Включаю {os.path.splitext(os.path.basename(f))[0]}."
    except Exception: return None

def stop_music() -> str:
    for p in ("wmplayer","vlc","spotify","groove","musicbee"):
        try: subprocess.run(["powershell.exe",f"Stop-Process -Name {p} -ErrorAction SilentlyContinue"],check=False)
        except Exception: pass
    speak_resp("stop_music"); return ""

def set_timer(seconds: int) -> str:
    def _t(): time.sleep(seconds); speak("Таймер сработал!")
    threading.Thread(target=_t, daemon=True).start()
    m, s = divmod(seconds, 60)
    return f"Таймер на {m} мин." if m else f"Таймер на {s} сек."

def set_alarm(hour: int, minute: int) -> str:
    global alarm_thread
    def _a():
        while True:
            n = datetime.now()
            if n.hour == hour and n.minute == minute:
                speak(f"Будильник! {hour:02d}:{minute:02d}!"); break
            time.sleep(20)
    alarm_thread = threading.Thread(target=_a, daemon=True)
    alarm_thread.start()
    return f"Будильник на {hour:02d}:{minute:02d}."

def stop_alarm()         -> str: global alarm_thread; alarm_thread = None; speak_resp("alarm_off"); return ""
def break_reminder_on()  -> str:
    global break_reminder_active
    if break_reminder_active: speak_resp("break_already"); return ""
    break_reminder_active = True
    def _r():
        while break_reminder_active:
            time.sleep(1800)
            if break_reminder_active: speak("Ты работаешь 30 минут. Перерыв!")
    threading.Thread(target=_r, daemon=True).start()
    speak_resp("break_on"); return ""

def break_reminder_off() -> str:
    global break_reminder_active; break_reminder_active = False
    speak_resp("break_off"); return ""

def shutdown()        -> str: os.system("shutdown /s /t 10"); speak_resp("shutdown"); return ""
def restart()         -> str: os.system("shutdown /r /t 10"); speak_resp("restart"); return ""
def sleep_pc()        -> str: os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0"); speak_resp("sleep"); return ""
def cancel_shutdown() -> str: os.system("shutdown /a"); speak_resp("cancel_shutdown"); return ""

def remind_me(minutes: int = 5, text: str = "Напоминание!") -> str:
    def _r(): time.sleep(minutes*60); speak(f"Напоминаю: {text}")
    threading.Thread(target=_r, daemon=True).start()
    return f"Напомню через {minutes} мин."

def calculate(expression: str = "") -> str:
    reps = {"плюс":"+","минус":"-","умножить на":"*","умножить":"*",
            "разделить на":"/","разделить":"/","в степени":"**"}
    expr = expression.lower()
    for w, op in reps.items(): expr = expr.replace(w, op)
    safe = "".join(c for c in expr if c in "0123456789+-*/(). ").strip()
    if not safe: return None
    try:
        result = eval(safe)
        if isinstance(result, float) and result == int(result): result = int(result)
        return f"{expression} равно {result}."
    except Exception: return None

def show_commands() -> str:
    categories = {
        "Время":        ["который час", "какая дата"],
        "Громкость":    ["громче", "тише", "выключи звук", "включи звук"],
        "Яркость":      ["ярче", "темнее"],
        "Приложения":   ["открой хром", "открой телеграм", "открой ворд", "..."],
        "Папки":        ["открой загрузки", "открой документы", "..."],
        "Система":      ["скриншот", "батарея", "загрузка процессора"],
        "WiFi":         ["включи вайфай", "выключи вайфай"],
        "Задачи":       ["добавь задачу", "список задач", "очисти задачи"],
        "Режимы":       ["ночной режим", "режим презентации", "утренний режим"],
        "Питание":      ["выключи компьютер", "перезагрузи", "спящий режим"],
        "ИИ":           ["любой вопрос свободным текстом"],
    }
    print("\n  ┌─────────────────────────────────────┐")
    print("  │         КОМАНДЫ ЛОРЫ                │")
    print("  ├─────────────────────────────────────┤")
    for cat, cmds in categories.items():
        print(f"  │  {cat:<12} {', '.join(cmds)[:24]:<24} │")
    print("  └─────────────────────────────────────┘\n")
    return "Список команд в консоли."


def break_code():
    speak_resp("farewell") or speak("До встречи!")
    time.sleep(1); exit()


# ─── РЕЖИМЫ ─────────────────────────────────────────────────────────────────

def mode_night() -> str:
    brightness_down(); brightness_down(); brightness_down()
    try:
        subprocess.run(["powershell","-Command",
            "Set-ItemProperty -Path 'HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion"
            "\\Notifications\\Settings' -Name 'NOC_GLOBAL_SETTING_TOASTS_ENABLED' -Value 0 -Force"],
            check=False, capture_output=True)
    except Exception: pass
    speak_resp("mode_night"); return ""

def mode_presentation() -> str:
    brightness_up(); brightness_up()
    try: subprocess.Popen(APP_PATHS.get("powerpoint",""), shell=True)
    except Exception: pass
    time.sleep(0.5); open_app("телеграм")
    time.sleep(0.5); open_folder("загрузки")
    time.sleep(0.5); webbrowser.open("https://gamma.app")
    speak_resp("mode_presentation"); return ""

def mode_morning() -> str:
    brightness_up(); brightness_up(); brightness_up()
    open_app("хром"); time.sleep(0.3); open_app("телеграм")
    mo = ["января","февраля","марта","апреля","мая","июня",
          "июля","августа","сентября","октября","ноября","декабря"]
    n  = datetime.now()
    wd = ["понедельник","вторник","среда","четверг","пятница","суббота","воскресенье"][n.weekday()]
    return f"Доброе утро! Сейчас {n.strftime('%H:%M')}, {wd} {n.day} {mo[n.month-1]}."

def _first_greeting() -> str:
    mo = ["января","февраля","марта","апреля","мая","июня",
          "июля","августа","сентября","октября","ноября","декабря"]
    n  = datetime.now()
    wd = ["понедельник","вторник","среда","четверг","пятница","суббота","воскресенье"][n.weekday()]
    h  = n.hour
    gr = "Доброй ночи" if h < 6 else "Доброе утро" if h < 12 else "Добрый день" if h < 18 else "Добрый вечер"
    return f"{gr}! Сейчас {n.strftime('%H:%M')}, {wd} {n.day} {mo[n.month-1]}. Чем могу помочь?"


# ─── EXECUTE ────────────────────────────────────────────────────────────────

def execute_command(cmd: str, query: str = "") -> str | None:
    if cmd.startswith("open_folder:"): return open_folder(cmd.split(":",1)[1])
    if cmd.startswith("open_app:"):    return open_app(cmd.split(":",1)[1])
    if cmd.startswith("close_app:"):   return close_app(cmd.split(":",1)[1])
    if cmd == "ping":
        keys = [k for k in _resp_cache if k.startswith("ping_")]
        if keys:
            chosen = random.choice(keys)
            _log_lora(ALL_RESPONSES[chosen])
            _play_file(_resp_cache[chosen])
        return ""
    if cmd == "mute":
        global is_muted; is_muted = True; stop_speech(); return ""
    if cmd == "unmute":
        is_muted = False
        keys = [k for k in _resp_cache if k.startswith("unmute_")]
        if keys:
            chosen = random.choice(keys)
            _log_lora(ALL_RESPONSES[chosen])
            _play_file(_resp_cache[chosen])
        return ""
    if cmd == "set_timer":
        secs = SlotExtractor.time_to_seconds(query)
        return set_timer(secs) if secs else "На сколько таймер?"
    if cmd == "set_alarm":
        m = re.search(r'(\d{1,2})[:. ](\d{2})', query)
        if m: return set_alarm(int(m.group(1)), int(m.group(2)))
        nums = re.findall(r'\d+', query)
        if len(nums) >= 2: return set_alarm(int(nums[0]), int(nums[1]))
        if len(nums) == 1: return set_alarm(int(nums[0]), 0)
        return None
    if cmd == "calculate": return calculate(query)

    dispatch = {
        "get_time":get_time,"get_date":get_date,
        "volume_up":volume_up,"volume_down":volume_down,
        "volume_max":volume_max,"volume_min":volume_min,
        "sound_off":sound_off,"sound_on":sound_on,
        "brightness_up":brightness_up,"brightness_down":brightness_down,
        "wifi_toggle":wifi_toggle,"wifi_toggle_on":wifi_toggle_on,"wifi_toggle_off":wifi_toggle_off,
        "screenshot":screenshot,"clipboard_copy":clipboard_copy,
        "clipboard_paste":clipboard_paste,"clipboard_read":clipboard_read,
        "switch_window":switch_window,"window_minimize":window_minimize,
        "window_maximize":window_maximize,"window_close":window_close,
        "get_battery":get_battery,"get_cpu":get_cpu,
        "get_weather":get_weather,"play_music":play_music,"stop_music":stop_music,
        "create_task":create_task,"show_tasks":show_tasks,"clear_tasks":clear_tasks,
        "open_browser":open_browser,"close_browser":close_browser,
        "open_folder":open_folder,
        "shutdown":shutdown,"restart":restart,"sleep":sleep_pc,
        "cancel_shutdown":cancel_shutdown,"stop_alarm":stop_alarm,
        "break_reminder_on":break_reminder_on,"break_reminder_off":break_reminder_off,
        "mode_night":mode_night,"mode_presentation":mode_presentation,"mode_morning":mode_morning,
        "break_code":break_code,
        "show_commands":show_commands,
        "remind_me":lambda: remind_me(5, "Напоминание"),
    }
    fn = dispatch.get(cmd)
    return fn() if fn else None


# ─── PROCESS ────────────────────────────────────────────────────────────────

def _process(query: str, wake_listener=None):
    global is_muted
    if is_speaking: stop_speech()
    query = query.lower().strip()

    if any(w in query for w in STOP_TRIGGERS):
        break_code()

    # Ping — обращение к Лоре (точное или fuzzy: "лара", "лаура", "лёра" и т.д.)
    _is_ping = False
    if len(query.split()) <= 2:
        if "лора" in query or "lora" in query:
            _is_ping = True
        else:
            try:
                from rapidfuzz import fuzz
                for w in ["лора", "эй лора"]:
                    if fuzz.ratio(query, w) >= 70:
                        _is_ping = True
                        break
            except ImportError:
                pass

    if _is_ping:
        keys = [k for k in _resp_cache if k.startswith("ping_")]
        if keys:
            chosen = random.choice(keys)
            _log_lora(ALL_RESPONSES[chosen])
            _play_file(_resp_cache[chosen])
        return

    if any(w in query for w in UNMUTE_TRIGGERS):
        if is_muted:
            is_muted = False
            keys = [k for k in _resp_cache if k.startswith("unmute_")]
            if keys:
                chosen = random.choice(keys)
                _log_lora(ALL_RESPONSES[chosen])
                _play_file(_resp_cache[chosen])
        return

    if any(w in query for w in MUTE_TRIGGERS):
        is_muted = True; stop_speech(); return

    if is_muted: return

    # 1. Точное совпадение
    cmd = LOCAL_COMMANDS.get(query)

    # 2. Rapidfuzz — порог 80% чтобы избежать ложных срабатываний
    if not cmd:
        try:
            from rapidfuzz import process, fuzz
            match = process.extractOne(
                query, LOCAL_COMMANDS.keys(),
                scorer=fuzz.token_set_ratio, score_cutoff=80,
            )
            if match:
                cmd = LOCAL_COMMANDS[match[0]]
        except ImportError:
            pass

    # 3. Частичное совпадение только для длинных фраз
    if not cmd and len(query.split()) >= 3:
        for phrase, c in LOCAL_COMMANDS.items():
            if len(phrase.split()) >= 2 and phrase in query:
                cmd = c; break

    if cmd:
        result = execute_command(cmd, query)
        if result is None:
            play_unclear()
        elif result == "":
            pass
        else:
            speak(result)
    else:
        if AI_ENABLED and _vosk_listener:
            free_text  = _vosk_listener.listen_free(timeout=5)
            full_query = (query + " " + free_text).strip() if free_text else query
            ai_reply   = ask_ai(full_query)
            if ai_reply: speak(ai_reply)
            else: play_unclear()
        else:
            play_unclear()


# ─── MAIN ────────────────────────────────────────────────────────────────────

_exit_event = threading.Event()

def _keyboard_watcher():
    def on_key(e):
        if e.name == "esc":
            if is_speaking:
                stop_speech()
            else:
                _exit_event.set()
    keyboard.on_press(on_key)
    _exit_event.wait()
    keyboard.unhook_all()

def main():
    global listen_fn, is_muted, _vosk_listener, _first_activation

    os.makedirs("music", exist_ok=True)
    os.makedirs("sounds", exist_ok=True)
    if not os.path.exists("список дел.txt"):
        open("список дел.txt","w",encoding="utf-8").close()

    _generate_cache()
    _vol_init()

    vosk_listener  = VoskListener(list(LOCAL_COMMANDS.keys()))
    _vosk_listener = vosk_listener
    listen_fn      = lambda timeout=8: vosk_listener.listen(timeout)

    _wake_event = threading.Event()

    def _on_wake():
        _wake_event.set()

    wake_listener = WakeWordListener(vosk_listener._model, _on_wake)
    wake_listener.start()

    speak_resp("ready") or speak("Готова к работе.")
    print("\n  Говори 'Эй Лора'  |  Escape — выйти\n")

    kbd_thread = threading.Thread(target=_keyboard_watcher, daemon=True)
    kbd_thread.start()

    try:
        while not _exit_event.is_set():
            triggered = _wake_event.wait(timeout=0.2)
            if not triggered:
                continue
            _wake_event.clear()

            # Пауза wake listener пока обрабатываем команду
            wake_listener.pause()

            if is_muted:
                is_muted = False

            # Первое приветствие
            if _first_activation:
                _first_activation = False
                speak(_first_greeting())
            else:
                play_wake()

            _log_you("эй лора")
            last_active = time.time()

            time.sleep(0.15)
            cmd_text = vosk_listener.listen(timeout=6)

            if cmd_text:
                _log_you(cmd_text)
                _process(cmd_text, wake_listener)
                last_active = time.time()

            # Окно активности
            while (time.time() - last_active) < WINDOW_AFTER_AI:
                followup = vosk_listener.listen(timeout=2)
                if followup:
                    last_active = time.time()
                    _log_you(followup)
                    _process(followup, wake_listener)

            # Пока мут — ждём размута
            while is_muted:
                mute_text = vosk_listener.listen(timeout=5)
                if mute_text:
                    _process(mute_text, wake_listener)

            wake_listener.resume()

    except KeyboardInterrupt:
        print("\n  Завершение...")
    finally:
        wake_listener.stop()


if __name__ == "__main__":
    main()