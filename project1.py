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
import keyboard
from datetime import datetime
from dotenv import load_dotenv
import edge_tts
import pygame
import sounddevice as sd
import vosk

try:
    import pvporcupine
    from pvrecorder import PvRecorder
    PORCUPINE_AVAILABLE = True
except ImportError:
    PORCUPINE_AVAILABLE = False

load_dotenv()

PICOVOICE_KEY   = os.getenv("PICOVOICE_KEY", "")
PPN_PATH        = os.path.join(os.path.dirname(os.path.abspath(__file__)), "hey_lora.ppn")
EDGE_VOICE      = os.getenv("EDGE_VOICE", "ru-RU-SvetlanaNeural")
WEATHER_API_KEY = os.getenv("WEATHER_API_KEY", "")
MODEL_PATH      = os.path.join(os.path.dirname(os.path.abspath(__file__)), "model")

# Настройки из .env (можно менять без редактирования кода)
VOL_STEP_LOW    = float(os.getenv("VOL_STEP_LOW",  "0.10"))  # шаг до 50%
VOL_STEP_HIGH   = float(os.getenv("VOL_STEP_HIGH", "0.05"))  # шаг после 50%
WAKE_SENSITIVITY= float(os.getenv("WAKE_SENSITIVITY", "0.5")) # чувствительность wake word
WINDOW_AFTER_AI = int(os.getenv("WINDOW_AFTER_AI", "12"))     # окно активности (сек)

try:
    from groq import Groq
    _groq_client = Groq(api_key=os.getenv("GROQ_API_KEY")) if os.getenv("GROQ_API_KEY") else None
    AI_ENABLED   = _groq_client is not None
except ImportError:
    _groq_client = None
    AI_ENABLED   = False

pygame.mixer.init()

is_muted              = False
is_speaking           = False
stop_speaking_event   = threading.Event()
alarm_thread          = None
break_reminder_active = False
_vosk_listener        = None

WAKE_PHRASES    = ["Слушаю.", "Да.", "Здесь."]
UNMUTE_TRIGGERS = ("размут", "включись", "продолжай", "проснись", "вернись", "активируйся", "слушай")
MUTE_TRIGGERS   = ("замолчи", "молчи", "тихо", "пауза", "не слушай", "заткнись",
                   "хватит", "подожди", "погоди", "тишина", "умолкни", "мут", "помолчи")
STOP_TRIGGERS   = ("завершить работу", "заверши работу", "выключись", "завершись",
                   "закройся", "выход", "пока", "до свидания", "отключись", "выключи себя",
                   "стоп")

CONFIDENCE_EXECUTE = 85
CONFIDENCE_ASK     = 60

APP_PATHS = {
    "telegram":    os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Roaming\Telegram Desktop\Telegram.exe"),
    "телеграм":    os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Roaming\Telegram Desktop\Telegram.exe"),
    "discord":     os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Local\Discord\Update.exe"),
    "дискорд":     os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Local\Discord\Update.exe"),
    "chrome":      r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    "хром":        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    "notepad":     "notepad.exe",    "блокнот":     "notepad.exe",
    "calculator":  "calc.exe",       "калькулятор": "calc.exe",
    "paint":       "mspaint.exe",
    "проводник":   "explorer.exe",   "explorer":    "explorer.exe",
    "word":        r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    "excel":       r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
    "powerpoint":  r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
    "поверпоинт":  r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
    "steam":       r"C:\Program Files (x86)\Steam\steam.exe",
    "happ":         r"C:\Program Files\FlyFrogLLC\Happ\Happ.exe",
    "хапп":         r"C:\Program Files\FlyFrogLLC\Happ\Happ.exe",
}
APP_NAMES_RU = {
    "телеграм":"Телеграм", "telegram":"Телеграм",
    "дискорд":"Дискорд",   "discord":"Дискорд",
    "хром":"Хром",         "chrome":"Хром",
    "блокнот":"Блокнот",   "notepad":"Блокнот",
    "калькулятор":"Калькулятор", "calculator":"Калькулятор",
    "проводник":"Проводник", "explorer":"Проводник",
    "стим":"Стим",         "steam":"Стим",
    "ворд":"Word",         "word":"Word",
    "эксель":"Excel",      "excel":"Excel",
    "поверпоинт":"PowerPoint", "powerpoint":"PowerPoint",
    "хапп":"Happ",         "happ":"Happ",
}

LOCAL_COMMANDS = {
    "который час":"get_time", "сколько времени":"get_time",
    "какое время":"get_time", "время":"get_time",
    "какое сегодня число":"get_date", "какая дата":"get_date",
    "день недели":"get_date", "дата":"get_date",
    "увеличь громкость":"volume_up", "громче":"volume_up",
    "сделай громче":"volume_up", "прибавь громкость":"volume_up",
    "уменьши громкость":"volume_down", "тише":"volume_down",
    "сделай тише":"volume_down", "убавь громкость":"volume_down",
    "выключи звук":"sound_off", "без звука":"sound_off",
    "включи звук":"sound_on", "верни звук":"sound_on",
    "максимальная громкость":"volume_max",
    "минимальная громкость":"volume_min",
    "увеличь яркость":"brightness_up", "ярче":"brightness_up",
    "уменьши яркость":"brightness_down", "темнее":"brightness_down",
    "сверни окно":"window_minimize", "сверни все":"window_minimize",
    "сверни всё":"window_minimize", "убери все окна":"window_minimize",
    "разверни окно":"window_maximize", "закрой окно":"window_close",
    "переключи окно":"switch_window", "альт таб":"switch_window",
    "скопируй":"clipboard_copy", "вставь":"clipboard_paste",
    "что в буфере":"clipboard_read",
    "скриншот":"screenshot", "сделай скриншот":"screenshot",
    "скопируй текст с экрана":"ocr_copy",
    "скопируй с экрана":"ocr_copy",
    "скопируй экран":"ocr_copy",
    "прочитай экран":"ocr_copy",
    "текст с экрана":"ocr_copy",
    "переведи экран":"ocr_translate",
    "переведи с экрана":"ocr_translate",
    "переведи текст с экрана":"ocr_translate",
    "перевод с экрана":"ocr_translate",
    "заряд батареи":"get_battery", "батарея":"get_battery",
    "загрузка процессора":"get_cpu", "нагрузка системы":"get_cpu",
    "сколько места на диске":"disk_space",
    "какой ip":"get_ip",
    "скорость интернета":"speedtest",
    "покажи процессы":"top_processes",
    "включи вайфай":"wifi_toggle_on", "выключи вайфай":"wifi_toggle_off",
    "вайфай":"wifi_toggle",
    "открой браузер":"open_browser", "закрой браузер":"close_browser",
    "открой телеграм":"open_app:телеграм", "запусти телеграм":"open_app:телеграм",
    "открой дискорд":"open_app:дискорд",   "запусти дискорд":"open_app:дискорд",
    "открой хром":"open_app:хром",         "запусти хром":"open_app:хром",
    "открой блокнот":"open_app:блокнот",   "открой калькулятор":"open_app:калькулятор",
    "открой проводник":"open_app:проводник","открой стим":"open_app:steam",
    "открой ворд":"open_app:word",         "открой эксель":"open_app:excel",
    "открой хапп":"open_app:хапп",         "запусти хапп":"open_app:хапп",
    "закрой хапп":"close_app:хапп",
    "включи музыку":"play_music", "играй музыку":"play_music",
    "выключи музыку":"stop_music", "стоп музыка":"stop_music",
    "следующий трек":"music_next", "предыдущий трек":"music_prev",
    "поставь на паузу":"music_pause", "продолжи музыку":"music_resume",
    "добавь задачу":"create_task", "новая задача":"create_task",
    "список задач":"show_tasks",   "мои задачи":"show_tasks",
    "очисти задачи":"clear_tasks",
    "открой загрузки":"open_folder:загрузки",
    "открой документы":"open_folder:документы",
    "открой рабочий стол":"open_folder:рабочий стол",
    "открой музыку":"open_folder:музыка",
    "открой видео":"open_folder:видео",
    "открой изображения":"open_folder:изображения",
    "таймер":"set_timer", "поставь таймер":"set_timer",
    "поставь будильник":"set_alarm", "отключи будильник":"stop_alarm",
    "выключи компьютер":"shutdown", "перезагрузи":"restart",
    "перезагрузка":"restart", "перезапусти код":"restart_script", "перезапуск кода":"restart_script", "спящий режим":"sleep",
    "отмени выключение":"cancel_shutdown",
    "включи перерывы":"break_reminder_on", "выключи перерывы":"break_reminder_off",
    "погода":"get_weather", "какая погода":"get_weather",
    "переведи":"translate",
    "диктуй":"dictate", "диктовка":"dictate",
    "найди файл":"find_file", "поиск файла":"find_file",
    "напомни":"remind_me",
    "посчитай":"calculate",
    "очисти буфер обмена":"clipboard_clear",
    "открой настройки":"open_settings",
    "блокировка экрана":"lock_screen",
    "сделай тёмную тему":"dark_mode",
    "ночной режим":"mode_night",
    "режим презентации":"mode_presentation",
    "утренний режим":"mode_morning",
    "доброе утро":"mode_morning",
    "замолчи":"mute", "мут":"mute", "молчи":"mute",
    "закрой телеграм":"close_app:телеграм", "выйди из телеграма":"close_app:телеграм",
    "закрой дискорд":"close_app:дискорд",   "выйди из дискорда":"close_app:дискорд",
    "закрой хром":"close_app:хром",          "закрой браузер хром":"close_app:хром",
    "закрой блокнот":"close_app:блокнот",
    "закрой ворд":"close_app:word",          "закрой эксель":"close_app:excel",
    "закрой стим":"close_app:steam",
    "размут":"unmute", "включись":"unmute", "слушай":"unmute",
    "эй лора":"ping", "хей лора":"ping", "hey lora":"ping",
    "какой сегодня праздник":"holiday", "что сегодня празднуют":"holiday",
    "день чего сегодня":"holiday", "какой праздник":"holiday",
    "праздник сегодня":"holiday",
    "расскажи факт":"fact_of_day", "интересный факт":"fact_of_day",
    "факт дня":"fact_of_day", "удиви меня":"fact_of_day",
    "расскажи анекдот":"tell_joke", "анекдот":"tell_joke",
    "смешной анекдот":"tell_joke", "пошути":"tell_joke",
    "дай совет":"daily_tip", "совет дня":"daily_tip",
    "совет на день":"daily_tip", "что посоветуешь":"daily_tip",
    "подбрось монетку":"coin_flip", "орёл или решка":"coin_flip",
    "монетка":"coin_flip", "монету подброси":"coin_flip",
    "подбрось монету":"coin_flip", "бросить монету":"coin_flip",
    "подброс монеты":"coin_flip", "кинь монету":"coin_flip",
}

CMD_NAMES = {
    "get_time":"узнать время", "get_date":"узнать дату",
    "volume_up":"громче", "volume_down":"тише",
    "volume_max":"максимальная громкость", "volume_min":"минимальная громкость",
    "sound_off":"выключить звук", "sound_on":"включить звук",
    "brightness_up":"ярче", "brightness_down":"темнее",
    "screenshot":"скриншот",
    "ocr_copy":"скопировать текст с экрана",
    "ocr_translate":"перевести текст с экрана",
    "clipboard_copy":"скопировать", "clipboard_paste":"вставить",
    "clipboard_read":"прочитать буфер", "clipboard_clear":"очистить буфер",
    "window_minimize":"свернуть окна", "window_maximize":"развернуть окно",
    "window_close":"закрыть окно", "switch_window":"переключить окно",
    "get_battery":"заряд батареи", "get_cpu":"загрузка процессора",
    "disk_space":"место на диске", "get_ip":"мой IP",
    "speedtest":"скорость интернета", "top_processes":"топ процессов",
    "wifi_toggle_on":"включить WiFi", "wifi_toggle_off":"выключить WiFi",
    "wifi_toggle":"переключить WiFi",
    "open_browser":"открыть браузер", "close_browser":"закрыть браузер",
    "play_music":"включить музыку", "stop_music":"выключить музыку",
    "music_next":"следующий трек", "music_prev":"предыдущий трек",
    "music_pause":"пауза", "music_resume":"продолжить музыку",
    "create_task":"добавить задачу", "show_tasks":"показать задачи",
    "clear_tasks":"очистить задачи",
    "set_timer":"таймер", "set_alarm":"будильник", "stop_alarm":"отключить будильник",
    "get_weather":"погода", "translate":"перевести фразу",
    "dictate":"диктовка", "find_file":"найти файл",
    "remind_me":"напоминание", "calculate":"калькулятор",
    "shutdown":"выключить компьютер", "restart":"перезагрузить",
    "sleep":"спящий режим", "cancel_shutdown":"отменить выключение",
    "break_reminder_on":"включить перерывы", "break_reminder_off":"выключить перерывы",
    "open_settings":"настройки", "lock_screen":"блокировать экран",
    "dark_mode":"тёмная тема",
    "mode_night":"ночной режим", "mode_morning":"утренний режим",
    "restart_script":"перезапуск кода",
    "mode_presentation":"режим презентации",
    "holiday":"праздник сегодня", "fact_of_day":"интересный факт",
    "tell_joke":"анекдот", "daily_tip":"совет дня", "coin_flip":"монетка",
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
    except Exception as e:
        print(f"  [!] commands.yaml: {e}")


# ─── TTS CACHE ───────────────────────────────────────────────────────────────

STATIC_RESPONSES = {
    "ready":"Готова к работе.",
    "unmuted":"Снова слушаю.",
    "unclear_0":"Не поняла.", "unclear_1":"Повтори пожалуйста.", "unclear_2":"Не расслышала.",
    "ping_0":"Да?", "ping_1":"Я здесь.", "ping_2":"Слушаю.",
    "vol_up":"Громче.", "vol_down":"Тише.",
    "vol_max":"Громкость максимальная.", "vol_min":"Громкость минимальная.",
    "sound_off":"Звук выключен.", "sound_on":"Звук включён.",
    "bright_up":"Ярче.", "bright_down":"Темнее.",
    "win_min":"Свёрнуто.", "win_max":"Развёрнуто.",
    "win_close":"Закрыто.", "win_switch":"Переключаю.",
    "copied":"Скопировано.", "pasted":"Вставлено.",
    "buf_empty":"Буфер пуст.", "buf_cleared":"Буфер очищен.",
    "screenshot":"Скриншот сохранён.",
    "ocr_done":"Текст скопирован в буфер.",
    "ocr_empty":"Текст на экране не найден.",
    "ocr_no_tesseract":"Tesseract не установлен. Установи через winget install UB-Mannheim.TesseractOCR",
    "wifi_on":"WiFi включён.", "wifi_off":"WiFi выключен.",
    "browser_closed":"Браузер закрыт.",
    "music_stopped":"Музыка остановлена.",
    "music_next":"Следующий трек.", "music_prev":"Предыдущий трек.",
    "music_pause":"Пауза.", "music_resume":"Продолжаю.",
    "tasks_empty":"Список пуст.", "tasks_cleared":"Список очищен.",
    "shutdown":"Выключаю через 10 секунд.",
    "restart":"Перезагружаю через 10 секунд.",
    "sleep":"Спящий режим.", "cancel_shutdown":"Выключение отменено.",
    "breaks_on":"Напоминания о перерывах включены.",
    "breaks_off":"Напоминания выключены.",
    "break_time":"Ты работаешь 30 минут. Пора сделать перерыв!",
    "dictated":"Напечатала.",
    "goodbye":"До встречи!",
    "locked":"Экран заблокирован.",
    "dark_mode":"Тёмная тема включена.",
    "settings_open":"Открываю настройки.",
    "mode_night":"Ночной режим. Яркость снижена.",
    "mode_morning":"Доброе утро! Хром и Телеграм открыты.",
    "mode_pres":"Режим презентации включён.",
    "stop_alarm":"Будильник отключён.",
    "confirm_yes":"Выполняю.",
    "confirm_no":"Отменила.",
    "wake_0":"Слушаю.", "wake_1":"Да.", "wake_2":"Здесь.",
    "coin_heads":"Орёл!", "coin_tails":"Решка!",
    "timer_done":"Таймер сработал!",
    "screenshot_done":"Скриншот сохранён.",
    "morning_done":"Утренний режим включён.",
    "pres_done":"Режим презентации включён. Телеграм, PowerPoint и Гамма открыты.",
    "night_done":"Ночной режим. Яркость снижена.",
}

_cache: dict = {}
_CACHE_DIR = os.path.join("sounds", "cache")


def _generate_cache():
    os.makedirs(_CACHE_DIR, exist_ok=True)

    async def _gen(text, path):
        await edge_tts.Communicate(text, EDGE_VOICE, rate="+10%").save(path)

    new_count = 0
    for key, text in STATIC_RESPONSES.items():
        path = os.path.join(_CACHE_DIR, f"{key}.mp3")
        if not os.path.exists(path):
            try:
                asyncio.run(_gen(text, path))
                new_count += 1
            except Exception:
                pass
        if os.path.exists(path):
            _cache[key] = path

    if new_count:
        print(f"  [tts] Сгенерировано {new_count} новых фраз")
    print(f"  [tts] Кэш: {len(_cache)} фраз")


def _play_file(path):
    try:
        pygame.mixer.music.load(path)
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy():
            pygame.time.Clock().tick(10)
        pygame.mixer.music.unload()
    except Exception:
        pass

_SFX_DIR = os.path.join("sounds", "sfx")

def _play_sfx(name):
    """Воспроизвести звуковой эффект из sounds/sfx/<name>.wav (или .mp3).
    Если файл не найден — молча пропустить."""
    for ext in ("wav", "mp3"):
        path = os.path.join(_SFX_DIR, f"{name}.{ext}")
        if os.path.exists(path):
            threading.Thread(target=_play_file, args=(path,), daemon=True).start()
            return


def play(key):
    if key in _cache:
        print(f"  lora  {STATIC_RESPONSES.get(key, key)}")
        _play_file(_cache[key])
    else:
        speak(STATIC_RESPONSES.get(key, key))


def play_random(prefix):
    keys = [k for k in _cache if k.startswith(prefix)]
    if keys:
        chosen = random.choice(keys)
        print(f"  lora  {STATIC_RESPONSES.get(chosen, chosen)}")
        _play_file(_cache[chosen])
    else:
        speak(random.choice(WAKE_PHRASES))


def speak(text):
    global is_speaking
    if not text:
        return
    print(f"  lora  {text}")

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
        except Exception:
            pass
        finally:
            is_speaking = False
            if tmp:
                try:
                    os.unlink(tmp)
                except Exception:
                    pass

    _worker()


def stop_speech():
    global is_speaking
    stop_speaking_event.set()
    try:
        pygame.mixer.music.stop()
    except Exception:
        pass
    is_speaking = False


# ─── VOSK ────────────────────────────────────────────────────────────────────

class VoskListener:
    SAMPLE_RATE = 16000
    BLOCK_SIZE  = 4000

    def __init__(self, phrases):
        if not os.path.exists(MODEL_PATH):
            raise RuntimeError(f"Модель Vosk не найдена: {MODEL_PATH}")
        vosk.SetLogLevel(-1)
        self._model    = vosk.Model(MODEL_PATH)
        grammar        = json.dumps(phrases + ["[unk]"], ensure_ascii=False)
        self._rec      = vosk.KaldiRecognizer(self._model, self.SAMPLE_RATE, grammar)
        self._rec_free = vosk.KaldiRecognizer(self._model, self.SAMPLE_RATE)
        print(f"  [vosk] Готов, фраз: {len(phrases)}")

    def listen(self, timeout=8):
        self._rec.Reset()
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
                if self._rec.AcceptWaveform(data):
                    text = json.loads(self._rec.Result()).get("text", "").strip()
                    if text and text != "[unk]":
                        return text

        text = json.loads(self._rec.FinalResult()).get("text", "").strip()
        return text if text and text != "[unk]" else None

    def listen_free(self, timeout=8):
        self._rec_free.Reset()
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
                if self._rec_free.AcceptWaveform(data):
                    text = json.loads(self._rec_free.Result()).get("text", "").strip()
                    if text:
                        return text

        text = json.loads(self._rec_free.FinalResult()).get("text", "").strip()
        return text if text else None


# ─── ИИ ──────────────────────────────────────────────────────────────────────

AI_SYSTEM_PROMPT = (
    "Тебя зовут Лора — голосовой ассистент. Общаешься как друг — просто, с лёгким юмором. "
    "Отвечай на русском. Простой вопрос — 1 предложение. Только текст."
)

def ask_ai(query):
    if not AI_ENABLED or not _groq_client:
        return ""
    try:
        r = _groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile", max_tokens=150, timeout=8,
            messages=[
                {"role": "system", "content": AI_SYSTEM_PROMPT},
                {"role": "user",   "content": f"[{datetime.now().strftime('%H:%M')}] {query}"},
            ],
        )
        return r.choices[0].message.content.strip()
    except Exception as e:
        return f"Ошибка: {e}"


# ─── OCR ─────────────────────────────────────────────────────────────────────

def _ocr_grab_text():
    """
    Делает скриншот В ПАМЯТИ (файл не сохраняется),
    читает текст через pytesseract.
    Возвращает: строку текста, None, или "no_tesseract".
    """
    try:
        import pyautogui
        import pytesseract

        img  = pyautogui.screenshot()   # PIL Image в памяти, без сохранения
        text = pytesseract.image_to_string(img, lang="rus+eng").strip()
        text = re.sub(r'[^\w\s.,!?:;\-\u2014\u00ab\u00bb"\'\(\)\n]', '', text)
        text = re.sub(r'\n{3,}', '\n\n', text).strip()
        return text if text else None

    except ImportError:
        return "no_tesseract"
    except Exception:
        return None


def ocr_copy():
    """Копирует весь текст с экрана в буфер обмена."""
    text = _ocr_grab_text()
    if text == "no_tesseract":
        return "ocr_no_tesseract"
    if not text:
        return "ocr_empty"
    try:
        import pyperclip
        pyperclip.copy(text)
        preview = text[:60].replace('\n', ' ')
        return f"Скопировал {len(text)} символов. Начало: {preview}{'…' if len(text) > 60 else ''}"
    except Exception:
        return "ocr_empty"


def ocr_translate():
    """Читает текст с экрана и переводит на русский."""
    text = _ocr_grab_text()
    if text == "no_tesseract":
        return "ocr_no_tesseract"
    if not text:
        return "ocr_empty"

    chunk = text[:500]

    if AI_ENABLED:
        try:
            r = _groq_client.chat.completions.create(
                model="llama-3.3-70b-versatile", max_tokens=300, timeout=10,
                messages=[
                    {"role": "system", "content":
                        "Переведи следующий текст на русский язык. "
                        "Верни только перевод, без пояснений."},
                    {"role": "user", "content": chunk},
                ],
            )
            return r.choices[0].message.content.strip()
        except Exception:
            pass

    try:
        from translate import Translator
        sentences = re.split(r'(?<=[.!?])\s+', chunk)[:5]
        t = Translator(from_lang="en", to_lang="ru")
        translated = [t.translate(s.strip()) for s in sentences if s.strip()]
        return " ".join(translated) if translated else "Не удалось перевести."
    except Exception:
        return "Не удалось перевести."


# ─── СЛОТ-ЭКСТРАКТОР ─────────────────────────────────────────────────────────

class SlotExtractor:
    @staticmethod
    def time_to_seconds(text):
        text = text.lower()
        if any(x in text for x in ["полчаса", "пол часа"]):
            return 1800
        total = 0
        for pat, mul in [(r'(\d+)\s*час', 3600), (r'(\d+)\s*мин', 60), (r'(\d+)\s*сек', 1)]:
            for m in re.findall(pat, text):
                total += int(m) * mul
        if total == 0:
            nums = re.findall(r'\d+', text)
            if nums:
                total = int(nums[0]) * 60
        return total if total > 0 else None


# ─── КОМАНДЫ ─────────────────────────────────────────────────────────────────

listen_fn = None

_vol_obj = None

def _vol():
    return _vol_obj

def _vol_init():
    global _vol_obj
    try:
        from pycaw.pycaw import AudioUtilities
        _vol_obj = AudioUtilities.GetSpeakers().EndpointVolume
        print(f"  [pycaw] Громкость: {int(round(_vol_obj.GetMasterVolumeLevelScalar()*100))}%")
    except Exception as e:
        print(f"  [!] pycaw недоступен: {e}")

def _wifi_status():
    try:
        r = subprocess.run(["netsh","interface","show","interface","name=Wi-Fi"],
                           capture_output=True, text=True, timeout=3)
        return "connect" in r.stdout.lower()
    except Exception:
        return False

def get_time():
    return f"Сейчас {datetime.now().strftime('%H:%M')}."

def get_date():
    mo = ["января","февраля","марта","апреля","мая","июня",
          "июля","августа","сентября","октября","ноября","декабря"]
    n  = datetime.now()
    wd = ["понедельник","вторник","среда","четверг","пятница","суббота","воскресенье"][n.weekday()]
    return f"Сегодня {n.day} {mo[n.month-1]} {n.year}, {wd}."

def _send_vol_key(key_code):
    """Эмулирует медиаклавишу — Windows показывает плашку громкости."""
    os.system(f"powershell.exe -WindowStyle Hidden (new-object -com wscript.shell).SendKeys([char]{key_code})")

def volume_up():
    v = _vol()
    if v:
        cur = v.GetMasterVolumeLevelScalar()
        step = VOL_STEP_HIGH if cur >= 0.5 else VOL_STEP_LOW
        new = min(1.0, cur + step)
        v.SetMasterVolumeLevelScalar(new, None)
    _send_vol_key(175)
    return "vol_up"

def volume_down():
    v = _vol()
    if v:
        cur = v.GetMasterVolumeLevelScalar()
        step = VOL_STEP_HIGH if cur > 0.5 else VOL_STEP_LOW
        new = max(0.0, cur - step)
        v.SetMasterVolumeLevelScalar(new, None)
    _send_vol_key(174)
    return "vol_down"

def volume_max():
    v = _vol()
    if v: v.SetMasterVolumeLevelScalar(1.0, None)
    _send_vol_key(175)

def volume_min():
    v = _vol()
    if v: v.SetMasterVolumeLevelScalar(0.0, None)
    _send_vol_key(174)

def sound_off():
    v = _vol()
    if v: v.SetMute(1, None)
    _send_vol_key(173)  # 173 = VK_VOLUME_MUTE

def sound_on():
    v = _vol()
    if v: v.SetMute(0, None)
    _send_vol_key(173)

def brightness_up():
    try:
        import screen_brightness_control as sbc
        new = min(100, sbc.get_brightness()[0] + 20)
        sbc.set_brightness(new)
        return f"Яркость {new}%."
    except Exception:
        os.system("powershell (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods)"
                  ".WmiSetBrightness(1,[math]::Min(100,((Get-WmiObject -Namespace root/WMI "
                  "-Class WmiMonitorBrightness).CurrentBrightness)+20))")
        return "bright_up"

def brightness_down():
    try:
        import screen_brightness_control as sbc
        new = max(0, sbc.get_brightness()[0] - 20)
        sbc.set_brightness(new)
        return f"Яркость {new}%."
    except Exception:
        os.system("powershell (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods)"
                  ".WmiSetBrightness(1,[math]::Max(0,((Get-WmiObject -Namespace root/WMI "
                  "-Class WmiMonitorBrightness).CurrentBrightness)-20))")
        return "bright_down"

def wifi_toggle_on():
    os.system('netsh interface set interface "Wi-Fi" enabled')

def wifi_toggle_off():
    os.system('netsh interface set interface "Wi-Fi" disabled')

def wifi_toggle():
    if _wifi_status(): os.system('netsh interface set interface "Wi-Fi" disabled')
    else:              os.system('netsh interface set interface "Wi-Fi" enabled')

def screenshot():
    try:
        import pyautogui
        _play_sfx("camera")
        pyautogui.screenshot(f"screenshot_{int(time.time())}.png")
        play("screenshot_done")
    except Exception:
        pass

def clipboard_copy():
    import pyautogui; pyautogui.hotkey("ctrl", "c")

def clipboard_paste():
    import pyautogui; pyautogui.hotkey("ctrl", "v")

def clipboard_clear():
    try:
        import pyperclip; pyperclip.copy("")
    except Exception:
        pass

def clipboard_read():
    try:
        import pyperclip; t = pyperclip.paste()
        return f"В буфере: {t[:200]}." if t else "buf_empty"
    except Exception:
        return "buf_empty"

def switch_window():
    import pyautogui; pyautogui.hotkey("alt", "tab")

def window_minimize():
    import pyautogui; pyautogui.hotkey("win", "d")

def window_maximize():
    import pyautogui; pyautogui.hotkey("win", "up")

def window_close():
    import pyautogui; pyautogui.hotkey("alt", "f4")

def get_battery():
    try:
        import psutil; b = psutil.sensors_battery()
        if b is None: return "Батарея не найдена."
        return f"Заряд {int(b.percent)}%, {'заряжается' if b.power_plugged else 'на батарее'}."
    except Exception:
        return "Не удалось."

def get_cpu():
    try:
        import psutil
        return f"Процессор {psutil.cpu_percent(interval=1)}%, память {psutil.virtual_memory().percent}%."
    except Exception:
        return "Не удалось."

def get_ip():
    try:
        import urllib.request
        ip = urllib.request.urlopen("https://api.ipify.org", timeout=4).read().decode()
        return f"Твой IP: {ip}."
    except Exception:
        return "Не удалось получить IP."

def disk_space():
    try:
        import shutil
        total, used, free = shutil.disk_usage("/")
        return f"Диск C: занято {used//2**30} из {total//2**30} ГБ, свободно {free//2**30} ГБ."
    except Exception:
        return "Не удалось."

def top_processes():
    try:
        import psutil
        procs = sorted(psutil.process_iter(['name','cpu_percent']),
                       key=lambda p: p.info['cpu_percent'], reverse=True)[:3]
        names = ", ".join(p.info['name'] for p in procs if p.info['cpu_percent'] > 0)
        return f"Топ: {names}." if names else "Все процессы в норме."
    except Exception:
        return "Не удалось."

def speedtest():
    try:
        import urllib.request
        start = time.time()
        urllib.request.urlopen("https://www.google.com", timeout=5)
        return f"Пинг до Google: {int((time.time()-start)*1000)} мс."
    except Exception:
        return "Нет интернета."

def lock_screen():
    os.system("rundll32.exe user32.dll,LockWorkStation")

def dark_mode():
    for prop in ("AppsUseLightTheme", "SystemUsesLightTheme"):
        subprocess.run(["powershell", "-Command",
            f"Set-ItemProperty -Path 'HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion"
            f"\\Themes\\Personalize' -Name '{prop}' -Value 0"],
            check=False, capture_output=True)

def open_settings():
    subprocess.Popen("ms-settings:", shell=True)

def music_next():
    import pyautogui; pyautogui.hotkey("ctrl", "right")

def music_prev():
    import pyautogui; pyautogui.hotkey("ctrl", "left")

def music_pause():
    import pyautogui; pyautogui.press("space")

def music_resume():
    import pyautogui; pyautogui.press("space")

def open_app(name):
    path    = APP_PATHS.get(name.lower(), name)
    ru_name = APP_NAMES_RU.get(name.lower(), name.capitalize())
    try:
        subprocess.Popen(path); return f"{ru_name} открыт."
    except Exception:
        try:
            subprocess.Popen(path, shell=True); return f"{ru_name} открыт."
        except Exception:
            return None

def close_app(name):
    process_map = {
        "телеграм":"telegram", "telegram":"telegram",
        "дискорд":"discord",   "discord":"discord",
        "хром":"chrome",       "chrome":"chrome",
        "блокнот":"notepad",   "notepad":"notepad",
        "ворд":"winword",      "word":"winword",
        "эксель":"excel",      "excel":"excel",
        "стим":"steam",        "steam":"steam",
        "хапп":"Happ",         "happ":"Happ",
    }
    ru_name = APP_NAMES_RU.get(name.lower(), name.capitalize())
    proc = process_map.get(name.lower(), name)
    try:
        subprocess.run(
            ["powershell.exe", f"Stop-Process -Name {proc} -ErrorAction SilentlyContinue"],
            check=False)
        return f"{ru_name} закрыт."
    except Exception:
        return None


def open_browser():
    speak("Какой сайт?"); site = listen_fn()
    if not site: return None
    url = (f"https://{site}" if "." in site and " " not in site
           else f"https://www.google.com/search?q={site.replace(' ', '+')}")
    webbrowser.open(url)
    return f"Открываю {site}."

def open_folder(folder_name=""):
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
        return f"{folder_name.capitalize()} открыта."
    return None

def get_weather():
    speak("Какой город?"); city = listen_fn()
    if not city: return None
    if not WEATHER_API_KEY: return "Добавь WEATHER_API_KEY в .env."
    try:
        import urllib.request, urllib.parse
        url = (f"https://api.openweathermap.org/data/2.5/weather"
               f"?q={urllib.parse.quote(city)}&appid={WEATHER_API_KEY}&units=metric&lang=ru")
        with urllib.request.urlopen(url, timeout=5) as r:
            d = json.loads(r.read().decode())
        return (f"В {city}: {d['weather'][0]['description']}, "
                f"{round(d['main']['temp'])}°, ощущается {round(d['main']['feels_like'])}°.")
    except Exception:
        return "Не удалось получить погоду."

def translate():
    speak("Скажи фразу."); t = listen_fn()
    if not t: return None
    try:
        from translate import Translator
        return f"По-английски: {Translator(from_lang='ru', to_lang='en').translate(t)}"
    except Exception:
        if AI_ENABLED: return ask_ai(f"Переведи на английский: {t}")
        return "Не удалось."

def create_task():
    speak("Что добавить?"); t = listen_fn()
    if not t: return None
    with open("список дел.txt", "a", encoding="utf-8") as f:
        f.write(f"✅ {t}\n")
    return f"Добавила: {t}."

def show_tasks():
    try:
        with open("список дел.txt", encoding="utf-8") as f:
            tasks = f.read().strip()
        if not tasks: return "tasks_empty"
        lines = tasks.splitlines()
        return f"Задач {len(lines)}: " + "; ".join(l.replace("✅ ", "") for l in lines[:5]) + "."
    except Exception:
        return "tasks_empty"

def clear_tasks():
    open("список дел.txt", "w", encoding="utf-8").close()
    return "tasks_cleared"

def play_music():
    try:
        files = [f for f in os.listdir("music") if f.endswith((".mp3",".wav",".flac"))]
        if not files: return "В папке music нет файлов."
        f = os.path.join("music", random.choice(files))
        os.startfile(f)
        return f"Включаю {os.path.splitext(os.path.basename(f))[0]}."
    except Exception:
        return None

def stop_music():
    for p in ("wmplayer","vlc","spotify","groove","musicbee"):
        try:
            subprocess.run(["powershell.exe",
                f"Stop-Process -Name {p} -ErrorAction SilentlyContinue"], check=False)
        except Exception:
            pass

def set_timer(seconds):
    def _t(): time.sleep(seconds); _play_sfx("timer"); play("timer_done")
    threading.Thread(target=_t, daemon=True).start()
    m, s = divmod(seconds, 60)
    return f"Таймер на {m} мин." if m else f"Таймер на {s} сек."

def set_alarm(hour, minute):
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

def stop_alarm():
    global alarm_thread; alarm_thread = None

def break_reminder_on():
    global break_reminder_active
    if break_reminder_active: return
    break_reminder_active = True
    def _r():
        while break_reminder_active:
            time.sleep(1800)
            if break_reminder_active: play("break_time")
    threading.Thread(target=_r, daemon=True).start()

def break_reminder_off():
    global break_reminder_active; break_reminder_active = False

def dictate():
    speak("Говори — напечатаю."); t = listen_fn()
    if not t: return None
    try:
        import pyautogui, pyperclip; pyperclip.copy(t); pyautogui.hotkey("ctrl","v")
        return "dictated"
    except Exception:
        return None

def find_file():
    speak("Что ищем?"); name = listen_fn()
    if not name: return None
    try:
        r = subprocess.run(["powershell.exe",
            f"Get-ChildItem -Path $env:USERPROFILE -Recurse -Filter '*{name}*' "
            f"-ErrorAction SilentlyContinue | Select-Object -First 3 FullName | Format-Table -HideTableHeaders"],
            capture_output=True, text=True, timeout=10)
        found = [l.strip() for l in r.stdout.strip().splitlines() if l.strip()]
        return f"Нашла: {found[0]}." if found else f"{name} не найден."
    except Exception:
        return None

def remind_me(minutes, text="Напоминание!"):
    def _r(): time.sleep(minutes*60); speak(f"Напоминаю: {text}")
    threading.Thread(target=_r, daemon=True).start()
    return f"Напомню через {minutes} мин."

def calculate(expression=""):
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
    except Exception:
        return None

def _arm_cancel(seconds=15):
    """Активирует _cancel_event на N секунд — Escape отменяет команду."""
    _cancel_event.set()
    def _clear(): time.sleep(seconds); _cancel_event.clear()
    threading.Thread(target=_clear, daemon=True).start()

def shutdown():
    os.system("shutdown /s /t 15")
    _arm_cancel(15)
    return "Выключаю через 15 секунд. Escape — отменить."

def restart():
    os.system("shutdown /r /t 15")
    _arm_cancel(15)
    return "Перезагружаю через 15 секунд. Escape — отменить."

def sleep_pc():
    _arm_cancel(15)
    def _do(): time.sleep(15); os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
    threading.Thread(target=_do, daemon=True).start()
    return "Спящий режим через 15 секунд. Escape — отменить."

def restart_script():
    import sys
    speak("Перезапускаю.")
    os.execv(sys.executable, [sys.executable] + sys.argv)

def cancel_shutdown(): os.system("shutdown /a"); return "Выключение отменено."

def mode_night():
    brightness_down(); brightness_down(); brightness_down()
    play("night_done")

def mode_morning():
    brightness_up(); brightness_up()
    open_app("хром"); time.sleep(0.3); open_app("телеграм")
    play("morning_done")

def mode_presentation():
    open_app("телеграм")
    open_app("powerpoint")
    import webbrowser
    webbrowser.open("https://gamma.app")
    play("pres_done")

def break_code():
    play("goodbye")
    _exit_event.set()


# ─── НЕЧЁТКИЙ ПОИСК ──────────────────────────────────────────────────────────

def _find_command(query):
    """Возвращает (cmd, score 0-100)."""
    # Точное совпадение — всегда
    if query in LOCAL_COMMANDS:
        return LOCAL_COMMANDS[query], 100

    # Для коротких запросов (1-2 слова) — только точное, без частичного
    if len(query.split()) <= 2:
        return None, 0

    # Частичное: фраза из словаря содержится в запросе или наоборот (3+ слов)
    best_cmd, best_score = None, 0
    for phrase, cmd in LOCAL_COMMANDS.items():
        if phrase in query:
            # Чем длиннее совпавшая фраза — тем лучше
            score = int(len(phrase) / len(query) * 100)
            if score > best_score:
                best_cmd, best_score = cmd, score
        elif query in phrase:
            score = int(len(query) / len(phrase) * 85)
            if score > best_score:
                best_cmd, best_score = cmd, score

    return (best_cmd, best_score) if best_cmd else (None, 0)


# ─── EXECUTE ─────────────────────────────────────────────────────────────────

_CACHE_RESPONSES = {
    "volume_max":"vol_max",    "volume_min":"vol_min",
    "sound_off":"sound_off",   "sound_on":"sound_on",
    "screenshot":"screenshot",
    "clipboard_copy":"copied", "clipboard_paste":"pasted",
    "clipboard_clear":"buf_cleared",
    "switch_window":"win_switch",
    "window_minimize":"win_min","window_maximize":"win_max",
    "window_close":"win_close",
    "stop_music":"music_stopped",
    "music_next":"music_next", "music_prev":"music_prev",
    "music_pause":"music_pause","music_resume":"music_resume",
    "wifi_toggle_on":"wifi_on","wifi_toggle_off":"wifi_off","wifi_toggle":"wifi_on",
    "stop_alarm":"stop_alarm",
    "break_reminder_on":"breaks_on","break_reminder_off":"breaks_off",
    "shutdown":"shutdown","restart":"restart",
    "sleep":"sleep","cancel_shutdown":"cancel_shutdown",
    "lock_screen":"locked","dark_mode":"dark_mode",
    "open_settings":"settings_open",
    "mode_night":"mode_night","mode_morning":"mode_morning","mode_presentation":"mode_pres",
    "clear_tasks":"tasks_cleared",
}


def _ask_groq_short(prompt):
    """Короткий запрос к Groq, возвращает строку ответа."""
    if not AI_ENABLED:
        return "ИИ недоступен."
    try:
        resp = _groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=300,
            temperature=0.9,
        )
        return resp.choices[0].message.content.strip()
    except Exception as e:
        return f"Ошибка: {e}"


def holiday():
    today = datetime.now()
    prompt = (
        f"Сегодня {today.day} {today.strftime('%B')} {today.year} года. "
        "Перечисли 3–5 праздников, которые отмечаются именно сегодня — международные, "
        "профессиональные, народные, необычные. Отвечай по-русски, кратко, "
        "без вступлений — сразу список через запятую или точку с запятой."
    )
    return _ask_groq_short(prompt)


def fact_of_day():
    today = datetime.now()
    prompt = (
        f"Сегодня {today.day} {today.strftime('%B')}. "
        "Расскажи один удивительный, малоизвестный факт — о науке, истории, природе или технологиях. "
        "2–3 предложения, по-русски, без вступлений."
    )
    return _ask_groq_short(prompt)


def tell_joke():
    prompt = (
        "Расскажи короткий анекдот или шутку для подростков 16-18 лет. "
        "Современный, про школу, игры, интернет, учёбу или жизнь. "
        "Без пошлости, без бородатых советских анекдотов. "
        "Только сам анекдот, без вступлений."
    )
    return _ask_groq_short(prompt)


def daily_tip():
    today = datetime.now()
    prompt = (
        f"Дай один практичный совет на сегодня ({today.strftime('%A, %d %B')}). "
        "Это может быть совет по продуктивности, здоровью, общению или настроению. "
        "1–2 предложения, по-русски, без вступлений."
    )
    return _ask_groq_short(prompt)


def coin_flip():
    _play_sfx("coin")
    key = random.choice(["coin_heads", "coin_tails"])
    return key


def execute_command(cmd, query=""):
    global is_muted

    if cmd.startswith("open_folder:"): return open_folder(cmd.split(":",1)[1])
    if cmd.startswith("open_app:"):    return open_app(cmd.split(":",1)[1])
    if cmd.startswith("close_app:"):   return close_app(cmd.split(":",1)[1])
    if cmd == "ping":   play_random("ping"); return "_done_"
    if cmd == "mute":   is_muted = True; stop_speech(); return "_mute_"
    if cmd == "unmute": is_muted = False; return "unmuted"
    if cmd == "ocr_copy":      return ocr_copy()
    if cmd == "ocr_translate": return ocr_translate()

    if cmd == "set_timer":
        secs = SlotExtractor.time_to_seconds(query)
        return set_timer(secs) if secs else "На сколько таймер?"
    if cmd == "set_alarm":
        m = re.search(r'(\d{1,2})[:.\s](\d{2})', query)
        if m: return set_alarm(int(m.group(1)), int(m.group(2)))
        nums = re.findall(r'\d+', query)
        if len(nums) >= 2: return set_alarm(int(nums[0]), int(nums[1]))
        if len(nums) == 1: return set_alarm(int(nums[0]), 0)
        return None
    if cmd == "calculate": return calculate(query)
    if cmd == "remind_me":
        secs = SlotExtractor.time_to_seconds(query)
        return remind_me(secs//60 if secs else 5, query)

    str_cmds = {
        "volume_up":volume_up, "volume_down":volume_down,
        "brightness_up":brightness_up, "brightness_down":brightness_down,
        "clipboard_read":clipboard_read,
        "get_battery":get_battery, "get_cpu":get_cpu,
        "get_ip":get_ip, "disk_space":disk_space,
        "top_processes":top_processes, "speedtest":speedtest,
        "get_time":get_time, "get_date":get_date,
        "get_weather":get_weather, "translate":translate,
        "play_music":play_music,
        "create_task":create_task, "show_tasks":show_tasks, "clear_tasks":clear_tasks,
        "music_next":music_next, "music_prev":music_prev,
        "music_pause":music_pause, "music_resume":music_resume,
        "stop_music":stop_music,
        "open_browser":open_browser, "open_folder":open_folder,
        "find_file":find_file, "dictate":dictate,
        "holiday":holiday, "fact_of_day":fact_of_day,
        "tell_joke":tell_joke, "daily_tip":daily_tip,
        "coin_flip":coin_flip,
    }
    if cmd in str_cmds:
        return str_cmds[cmd]()

    void_cmds = {
        "volume_max":volume_max, "volume_min":volume_min,
        "sound_off":sound_off, "sound_on":sound_on,
        "screenshot":screenshot,
        "clipboard_copy":clipboard_copy, "clipboard_paste":clipboard_paste,
        "clipboard_clear":clipboard_clear,
        "switch_window":switch_window,
        "window_minimize":window_minimize, "window_maximize":window_maximize,
        "window_close":window_close,

        "wifi_toggle":wifi_toggle, "wifi_toggle_on":wifi_toggle_on,
        "wifi_toggle_off":wifi_toggle_off,
        "shutdown":shutdown, "restart":restart,
        "sleep":sleep_pc, "cancel_shutdown":cancel_shutdown,
        "stop_alarm":stop_alarm,
        "break_reminder_on":break_reminder_on, "break_reminder_off":break_reminder_off,

        "lock_screen":lock_screen, "dark_mode":dark_mode,
        "open_settings":open_settings,
        "mode_night":mode_night, "mode_morning":mode_morning,
        "mode_presentation":mode_presentation,
        "restart_script":restart_script,
        "break_code":break_code,
    }
    fn = void_cmds.get(cmd)
    if fn:
        fn()
        return _CACHE_RESPONSES.get(cmd)
    return None


def _respond(result):
    if result is None or result in ("_done_", "_mute_"):
        return
    if isinstance(result, str) and result in _cache:
        play(result)
    elif isinstance(result, str):
        speak(result)


# ─── ОБРАБОТЧИК КОМАНДЫ ───────────────────────────────────────────────────────

def _process(query):
    global is_muted
    query = query.lower().strip()

    # wake word в окне активности — реагируем как на обычное обращение и выходим
    if any(w in query for w in ("эй лора", "лора", "hey lora")):
        play_random("wake")
        return

    # "стоп" во время речи — только прерываем, не выходим
    if is_speaking:
        stop_speech()
        if query == "стоп":
            return

    # ── ФИКС: печатаем фразу ДО проверки стоп-триггеров ──
    print(f"  \033[1myou   {query}\033[0m")

    if any(w in query for w in STOP_TRIGGERS):
        break_code()

    if any(w in query for w in UNMUTE_TRIGGERS):
        if is_muted:
            is_muted = False
            play("unmuted")
        return

    if any(w in query for w in MUTE_TRIGGERS):
        is_muted = True
        stop_speech()
        print()
        print("  ┌─────────────────────────────────┐")
        print("  │  МУТ — Лора молчит              │")
        print("  │  Чтобы разбудить, скажи:        │")
        print("  │  'размут'  'включись'  'слушай' │")
        print("  │  'продолжай'  'проснись'        │")
        print("  └─────────────────────────────────┘")
        print()
        return

    if is_muted: return

    cmd, score = _find_command(query)
    print(f"  \033[3mconf  {score}%  →  {cmd}\033[0m")

    # Зона 1: уверенно → выполняем сразу
    if score >= CONFIDENCE_EXECUTE and cmd:
        result = execute_command(cmd, query)
        _respond(result)
        return

    # Зона 2: сомнение → переспрашиваем
    if CONFIDENCE_ASK <= score < CONFIDENCE_EXECUTE and cmd:
        cmd_name = CMD_NAMES.get(cmd, cmd)
        if cmd.startswith("open_app:"):
            app = cmd.split(":",1)[1]
            cmd_name = f"открыть {APP_NAMES_RU.get(app, app)}"
        elif cmd.startswith("open_folder:"):
            cmd_name = f"открыть папку {cmd.split(':',1)[1]}"

        speak(f"Ты имеешь в виду «{cmd_name}»?")
        confirm = _vosk_listener.listen(timeout=5) if _vosk_listener else None
        print(f"  \033[1myou   {confirm or '—'}\033[0m")

        if confirm and any(w in confirm.lower() for w in
                           ("да", "верно", "точно", "именно", "ага", "угу", "конечно")):
            play("confirm_yes")
            result = execute_command(cmd, query)
            _respond(result)
        else:
            play("confirm_no")
        return

    # Зона 3: не понял → ИИ или unclear
    if AI_ENABLED and _vosk_listener:
        free_text  = _vosk_listener.listen_free(timeout=5)
        full_query = (query + " " + free_text).strip() if free_text else query
        ai_reply   = ask_ai(full_query)
        if ai_reply: speak(ai_reply)
        else: play_random("unclear")
    else:
        play_random("unclear")


# ─── MAIN ─────────────────────────────────────────────────────────────────────

_exit_event = threading.Event()
_cancel_event = threading.Event()  # Escape отменяет выключение/сон/перезагрузку

def _keyboard_watcher():
    """Отдельный поток: Escape прерывает речь, отменяет выключение или завершает программу."""
    def on_key(e):
        if e.name == "esc":
            if _cancel_event.is_set():
                # Идёт обратный отсчёт выключения/сна/перезагрузки — отменяем
                _cancel_event.clear()
                os.system("shutdown /a")
                print("  [kbd] выключение отменено")
                speak("Отменила.")
            elif is_speaking:
                stop_speech()
                print("  [kbd] речь прервана")
            else:
                print("  [kbd] завершение...")
                _exit_event.set()
    keyboard.on_press(on_key)
    _exit_event.wait()
    keyboard.unhook_all()


def main():
    global listen_fn, _vosk_listener, is_muted

    os.makedirs("music", exist_ok=True)
    os.makedirs("sounds", exist_ok=True)
    if not os.path.exists("список дел.txt"):
        open("список дел.txt", "w", encoding="utf-8").close()

    _generate_cache()
    _vol_init()

    if not PICOVOICE_KEY:
        print("  [!] Добавь PICOVOICE_KEY в .env"); exit(1)
    if not os.path.exists(PPN_PATH):
        print(f"  [!] Файл не найден: {PPN_PATH}"); exit(1)

    vosk_listener  = VoskListener(list(LOCAL_COMMANDS.keys()))
    _vosk_listener = vosk_listener
    listen_fn      = lambda timeout=8: vosk_listener.listen(timeout)

    porcupine = pvporcupine.create(
        access_key=PICOVOICE_KEY,
        keyword_paths=[PPN_PATH],
        sensitivities=[WAKE_SENSITIVITY]
    )
    recorder = PvRecorder(device_index=-1, frame_length=porcupine.frame_length)
    recorder.start()

    play("ready")
    print("\n  Говори 'Эй Лора'  |  Escape — прервать речь или выйти\n")

    kbd_thread = threading.Thread(target=_keyboard_watcher, daemon=True)
    kbd_thread.start()

    try:
        while not _exit_event.is_set():
            pcm = recorder.read()
            if porcupine.process(pcm) >= 0:
                recorder.stop()
                print("  \033[1myou   эй лора\033[0m")
                if is_muted:
                    is_muted = False
                play_random("wake")
                last_active = time.time()

                time.sleep(0.2)
                cmd_text = vosk_listener.listen(timeout=6)
                recorder.start()

                if not cmd_text:
                    continue

                _process(cmd_text)
                last_active = time.time()

                # Если после команды вошли в мут — слушаем размут без wake word
                while is_muted:
                    mute_text = vosk_listener.listen(timeout=5)
                    if mute_text:
                        _process(mute_text)

                deadline = time.time() + WINDOW_AFTER_AI
                while time.time() < deadline and not _exit_event.is_set():
                    followup = vosk_listener.listen(timeout=1)
                    if followup:
                        last_active = time.time()
                        deadline = last_active + WINDOW_AFTER_AI
                        _process(followup)

    except KeyboardInterrupt:
        print("\n  Завершение...")
        _exit_event.set()
    finally:
        recorder.stop()
        recorder.delete()
        porcupine.delete()


if __name__ == "__main__":
    main()