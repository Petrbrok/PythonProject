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

# ── ИИ отключён временно, не удалять ────────────────────────────────────────
AI_ENABLED = False
try:
    from groq import Groq
    _groq_client = Groq(api_key=os.getenv("GROQ_API_KEY")) if os.getenv("GROQ_API_KEY") else None
except ImportError:
    _groq_client = None

pygame.mixer.init()

is_muted            = False
is_speaking         = False
stop_speaking_event = threading.Event()
alarm_thread        = None
break_reminder_active = False
WINDOW_AFTER_AI     = 12

WAKE_PHRASES  = ["Слушаю.", "Да.", "Здесь."]
MUTE_TRIGGERS = (
    "замолчи", "молчи", "тихо", "пауза", "не слушай",
    "заткнись", "хватит", "подожди", "погоди", "тишина",
    "умолкни", "мут", "стоп", "stop", "помолчи",
)
UNMUTE_TRIGGERS = (
    "размут", "включись", "продолжай", "проснись", "вернись",
    "активируйся", "слушай",
)
STOP_TRIGGERS = (
    "завершить работу", "заверши работу", "выключись", "завершись",
    "закройся", "выход", "пока", "до свидания", "отключись", "выключи себя",
)

CACHED_PHRASES = {
    "ready":     "Готова к работе",
    "unclear_0": "Не поняла",
    "unclear_1": "Повтори пожалуйста",
    "unclear_2": "Не расслышала",
}
_phrase_cache: dict[str, str] = {}

APP_PATHS = {
    "telegram":    os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Roaming\Telegram\Telegram.exe"),
    "телеграм":    os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Roaming\Telegram\Telegram.exe"),
    "discord":     os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Local\Discord\Update.exe"),
    "дискорд":     os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Local\Discord\Update.exe"),
    "spotify":     os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Roaming\Spotify\Spotify.exe"),
    "спотифай":    os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Roaming\Spotify\Spotify.exe"),
    "chrome":      r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    "хром":        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    "notepad":     "notepad.exe",   "блокнот":    "notepad.exe",
    "calculator":  "calc.exe",      "калькулятор":"calc.exe",
    "paint":       "mspaint.exe",
    "проводник":   "explorer.exe",  "explorer":   "explorer.exe",
    "word":        r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    "excel":       r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
    "obs":         r"C:\Program Files\obs-studio\bin\64bit\obs64.exe",
    "steam":       r"C:\Program Files (x86)\Steam\steam.exe",
    "vscode":      os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Local\Programs\Microsoft VS Code\Code.exe"),
    "код":         os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Local\Programs\Microsoft VS Code\Code.exe"),
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
    "obs":"OBS",
    "код":"VS Code",      "vscode":"VS Code",
    "ворд":"Word",        "word":"Word",
    "эксель":"Excel",     "excel":"Excel",
}

# ─── КОМАНДЫ ────────────────────────────────────────────────────────────────
LOCAL_COMMANDS = {
    # Время
    "который час":"get_time","сколько времени":"get_time",
    "какое время":"get_time","время":"get_time",
    "какое сегодня число":"get_date","какая дата":"get_date",
    "день недели":"get_date","дата":"get_date",
    # Громкость
    "увеличь громкость":"volume_up","громче":"volume_up",
    "сделай громче":"volume_up","прибавь громкость":"volume_up",
    "уменьши громкость":"volume_down","тише":"volume_down",
    "сделай тише":"volume_down","убавь громкость":"volume_down",
    "выключи звук":"sound_off","без звука":"sound_off",
    "включи звук":"sound_on","верни звук":"sound_on",
    "максимальная громкость":"volume_max",
    "поставь максимальную громкость":"volume_max",
    "минимальная громкость":"volume_min",
    # Яркость
    "увеличь яркость":"brightness_up","ярче":"brightness_up","сделай ярче":"brightness_up",
    "уменьши яркость":"brightness_down","темнее":"brightness_down","сделай темнее":"brightness_down",
    # Окна
    "сверни окно":"window_minimize","сверни все":"window_minimize",
    "сверни всё":"window_minimize","убери все окна":"window_minimize",
    "разверни окно":"window_maximize","закрой окно":"window_close",
    "переключи окно":"switch_window","альт таб":"switch_window","переключить окно":"switch_window",
    # Буфер
    "скопируй":"clipboard_copy","вставь":"clipboard_paste","что в буфере":"clipboard_read",
    # Скриншот
    "скриншот":"screenshot","сделай скриншот":"screenshot","снимок экрана":"screenshot",
    # Система
    "заряд батареи":"get_battery","сколько заряда":"get_battery","батарея":"get_battery",
    "загрузка процессора":"get_cpu","нагрузка системы":"get_cpu",
    # WiFi
    "включи вайфай":"wifi_toggle_on","включи вай фай":"wifi_toggle_on",
    "выключи вайфай":"wifi_toggle_off","выключи вай фай":"wifi_toggle_off",
    "вайфай":"wifi_toggle",
    # Браузер
    "открой браузер":"open_browser","закрой браузер":"close_browser",
    # Приложения
    "открой телеграм":"open_app:телеграм","запусти телеграм":"open_app:телеграм",
    "открой дискорд":"open_app:дискорд","запусти дискорд":"open_app:дискорд",
    "открой спотифай":"open_app:спотифай","запусти спотифай":"open_app:спотифай",
    "открой хром":"open_app:хром","запусти хром":"open_app:хром",
    "открой блокнот":"open_app:блокнот","открой калькулятор":"open_app:калькулятор",
    "открой проводник":"open_app:проводник","открой стим":"open_app:steam",
    "открой obs":"open_app:obs","открой код":"open_app:код",
    "открой ворд":"open_app:word","открой эксель":"open_app:excel",
    # Музыка
    "включи музыку":"play_music","играй музыку":"play_music",
    "выключи музыку":"stop_music","стоп музыка":"stop_music",
    # Задачи
    "добавь задачу":"create_task","новая задача":"create_task",
    "список задач":"show_tasks","мои задачи":"show_tasks","очисти задачи":"clear_tasks",
    # Папки
    "открой загрузки":"open_folder:загрузки",
    "открой документы":"open_folder:документы",
    "открой рабочий стол":"open_folder:рабочий стол",
    "открой музыку":"open_folder:музыка",
    "открой видео":"open_folder:видео",
    "открой изображения":"open_folder:изображения",
    # Таймер и будильник
    "таймер":"set_timer",
    "поставь таймер":"set_timer",
    "поставь будильник":"set_alarm",
    "будильник":"stop_alarm",
    "отключи будильник":"stop_alarm",
    # Питание
    "выключи компьютер":"shutdown","перезагрузи":"restart",
    "перезагрузка":"restart","спящий режим":"sleep","отмени выключение":"cancel_shutdown",
    # Перерывы
    "включи перерывы":"break_reminder_on","выключи перерывы":"break_reminder_off",
    # Мут/размут/пинг
    "замолчи":"mute","молчи":"mute","мут":"mute","тихо":"mute",
    "лора":"ping",
    "размут":"unmute","включись":"unmute","слушай":"unmute","продолжай":"unmute",
}

# Загружаем commands.yaml
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
    """
    Vosk с reusable KaldiRecognizer и жёсткой грамматикой.
    Reset() вместо создания нового объекта → ~0мс накладных расходов.
    """
    SAMPLE_RATE = 16000
    BLOCK_SIZE  = 4000   # 250мс

    def __init__(self, phrases: list[str]):
        if not os.path.exists(MODEL_PATH):
            raise RuntimeError(
                f"Модель Vosk не найдена: {MODEL_PATH}\n"
                f"Скачай vosk-model-ru-0.42 с https://alphacephei.com/vosk/models\n"
                f"и переименуй папку в 'model' рядом с project1.py"
            )
        vosk.SetLogLevel(-1)
        self._model    = vosk.Model(MODEL_PATH)
        grammar        = json.dumps(phrases + ["[unk]"], ensure_ascii=False)
        self._rec      = vosk.KaldiRecognizer(self._model, self.SAMPLE_RATE, grammar)
        self._rec.SetWords(True)
        print(f"  [vosk] Готов, фраз в грамматике: {len(phrases)}")

    def listen(self, timeout: float = 8) -> str | None:
        self._rec.Reset()
        q  = queue.Queue()
        t0 = time.time()

        def _cb(indata, frames, t, status):
            q.put(bytes(indata))

        with sd.RawInputStream(
            samplerate=self.SAMPLE_RATE,
            blocksize=self.BLOCK_SIZE,
            dtype="int16", channels=1,
            callback=_cb
        ):
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


# ─── TTS ────────────────────────────────────────────────────────────────────

def _generate_cache():
    os.makedirs(os.path.join("sounds","cache"), exist_ok=True)
    os.makedirs(os.path.join("sounds","wake"),  exist_ok=True)

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
        _log_lora(CACHED_PHRASES.get(key, key))
        _play_file(_phrase_cache[key])
    elif key in CACHED_PHRASES:
        speak(CACHED_PHRASES[key])


def play_unclear():
    keys = [k for k in _phrase_cache if k.startswith("unclear")]
    if keys:
        chosen = random.choice(keys)
        _log_lora(CACHED_PHRASES.get(chosen, chosen))
        _play_file(_phrase_cache[chosen])
    else:
        speak(random.choice(["Не поняла", "Повтори пожалуйста", "Не расслышала"]))


def play_wake():
    files = [f for f in os.listdir(os.path.join("sounds","wake")) if f.endswith(".mp3")] \
            if os.path.exists(os.path.join("sounds","wake")) else []
    if files:
        _play_file(os.path.join("sounds","wake", random.choice(files)))
    else:
        speak(random.choice(WAKE_PHRASES))


# ─── Silero TTS ─────────────────────────────────────────────────────────────
_silero_model  = None
_silero_sample = 48000

def _init_silero():
    global _silero_model, _silero_sample
    try:
        import torch
        device = torch.device("cpu")
        model, _ = torch.hub.load(
            repo_or_dir="snakers4/silero-models",
            model="silero_tts",
            language="ru",
            speaker="v4_ru",
            verbose=False,
        )
        model.to(device)
        _silero_model  = model
        _silero_sample = 48000
        print("  [tts] Silero загружен")
        return True
    except Exception as e:
        print(f"  [!] Silero недоступен: {e}")
        return False

def _speak_silero(text: str):
    """Синтез через Silero — офлайн ~100-200мс."""
    global is_speaking
    is_speaking = True
    stop_speaking_event.clear()
    try:
        import torch
        audio = _silero_model.apply_tts(
            text=text,
            speaker="baya",        # женский голос
            sample_rate=_silero_sample,
            put_accent=True,
            put_yo=True,
        )
        # audio — tensor float32, конвертируем в int16 для pygame
        import numpy as np
        audio_np = (audio.numpy() * 32767).astype("int16")
        # pygame stereo требует 2D массив (samples, 2)
        audio_np = np.column_stack([audio_np, audio_np])
        sound = pygame.sndarray.make_sound(audio_np)
        channel = sound.play()
        while channel.get_busy():
            if stop_speaking_event.is_set():
                channel.stop()
                break
            pygame.time.Clock().tick(10)
    except Exception:
        _speak_edge(text)   # fallback
    finally:
        is_speaking = False

def _speak_edge(text: str):
    """Fallback — edge-tts через интернет."""
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
            try: os.unlink(tmp)
            except Exception: pass

def speak(text: str):
    global is_speaking
    if not text: return
    _log_lora(text)

    def _worker():
        if _silero_model is not None:
            _speak_silero(text)
        else:
            _speak_edge(text)

    t = threading.Thread(target=_worker, daemon=True)
    t.start()
    t.join()


def stop_speech():
    global is_speaking
    stop_speaking_event.set()
    try: pygame.mixer.music.stop()
    except Exception: pass
    is_speaking = False


# ─── КОНСОЛЬ ────────────────────────────────────────────────────────────────

def _log_you(text: str):
    print(f"  you   {text}")

def _log_lora(text: str, ms: int | None = None):
    suffix = f"  [{ms}мс]" if ms is not None else ""
    print(f"  lora  {text}{suffix}")


# ─── ИИ (отключён) ──────────────────────────────────────────────────────────

AI_SYSTEM_PROMPT = (
    "Тебя зовут Лора — голосовой ассистент. Общаешься как друг — просто, с лёгким юмором. "
    "Отвечай на русском. Простой вопрос — 1 предложение. Только текст."
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


# ─── СЛОТ-ЭКСТРАКТОР (из новой версии) ──────────────────────────────────────

class SlotExtractor:
    @staticmethod
    def time_to_seconds(text: str) -> int | None:
        text = text.lower()
        if any(x in text for x in ["полчаса","пол часа"]): return 1800
        if re.search(r'(?<!\d\s)час(?!а?\s*\d)', text): return 3600
        total = 0
        for pat, mul in [(r'(\d+)\s*час',3600),(r'(\d+)\s*мин',60),(r'(\d+)\s*сек',1)]:
            for m in re.findall(pat, text):
                total += int(m) * mul
        if total == 0:
            nums = re.findall(r'\d+', text)
            if nums: total = int(nums[0]) * 60
        return total if total > 0 else None


# ─── КОМАНДЫ ────────────────────────────────────────────────────────────────

listen_fn = None   # устанавливается в main()

def _vol():
    try:
        from ctypes import POINTER, cast
        from comtypes import CLSCTX_ALL
        from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
        i = AudioUtilities.GetSpeakers().Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        return cast(i, POINTER(IAudioEndpointVolume))
    except Exception as e:
        return None

def _vol_init():
    """Инициализируем pycaw один раз при старте."""
    v = _vol()
    if v is None:
        print("  [!] pycaw недоступен — громкость через клавиши")
    else:
        pct = int(round(v.GetMasterVolumeLevelScalar() * 100))
        print(f"  [pycaw] Громкость: {pct}%")

def _get_vol_pct() -> int | None:
    v = _vol()
    return int(round(v.GetMasterVolumeLevelScalar() * 100)) if v else None

def _wifi_status() -> bool:
    try:
        r = subprocess.run(
            ["netsh","interface","show","interface","name=Wi-Fi"],
            capture_output=True, text=True, timeout=3)
        return "connect" in r.stdout.lower()
    except Exception:
        return False

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
        return f"Громкость {int(round(new*100))}%."
    try:
        import pyautogui
        for _ in range(5): pyautogui.press("volumeup")
        pct = _get_vol_pct()
        return f"Громкость {pct}%." if pct else "Громкость увеличена."
    except Exception: return "Громкость увеличена."

def volume_down() -> str:
    v = _vol()
    if v:
        new = max(0.0, v.GetMasterVolumeLevelScalar() - 0.15)
        v.SetMasterVolumeLevelScalar(new, None)
        return f"Громкость {int(round(new*100))}%."
    try:
        import pyautogui
        for _ in range(5): pyautogui.press("volumedown")
        pct = _get_vol_pct()
        return f"Громкость {pct}%." if pct else "Громкость уменьшена."
    except Exception: return "Громкость уменьшена."

def volume_max() -> str:
    v = _vol()
    if v: v.SetMasterVolumeLevelScalar(1.0, None)
    return "Громкость максимальная."

def volume_min() -> str:
    v = _vol()
    if v: v.SetMasterVolumeLevelScalar(0.0, None)
    return "Громкость минимальная."

def sound_off() -> str:
    v = _vol()
    if v: v.SetMute(1, None)
    else: os.system("powershell.exe (new-object -com wscript.shell).SendKeys([char]173)")
    return "Звук выключен."

def sound_on() -> str:
    v = _vol()
    if v: v.SetMute(0, None)
    else: os.system("powershell.exe (new-object -com wscript.shell).SendKeys([char]173)")
    return "Звук включён."

def brightness_up() -> str:
    try:
        import screen_brightness_control as sbc
        new = min(100, sbc.get_brightness()[0] + 20)
        sbc.set_brightness(new); return f"Яркость {new}%."
    except Exception:
        os.system("powershell (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods)"
                  ".WmiSetBrightness(1,[math]::Min(100,((Get-WmiObject -Namespace root/WMI "
                  "-Class WmiMonitorBrightness).CurrentBrightness)+20))")
        return "Яркость увеличена."

def brightness_down() -> str:
    try:
        import screen_brightness_control as sbc
        new = max(0, sbc.get_brightness()[0] - 20)
        sbc.set_brightness(new); return f"Яркость {new}%."
    except Exception:
        os.system("powershell (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods)"
                  ".WmiSetBrightness(1,[math]::Max(0,((Get-WmiObject -Namespace root/WMI "
                  "-Class WmiMonitorBrightness).CurrentBrightness)-20))")
        return "Яркость уменьшена."

def wifi_toggle_on()  -> str:
    os.system('netsh interface set interface "Wi-Fi" enabled');  return "WiFi включён."
def wifi_toggle_off() -> str:
    os.system('netsh interface set interface "Wi-Fi" disabled'); return "WiFi выключен."
def wifi_toggle()     -> str:
    if _wifi_status():
        os.system('netsh interface set interface "Wi-Fi" disabled'); return "WiFi выключен."
    os.system('netsh interface set interface "Wi-Fi" enabled'); return "WiFi включён."

def screenshot() -> str:
    try:
        import pyautogui
        pyautogui.screenshot(f"screenshot_{int(time.time())}.png")
        return "Скриншот сохранён."
    except Exception: return None

def clipboard_copy()  -> str:
    import pyautogui; pyautogui.hotkey("ctrl","c"); return "Скопировано."
def clipboard_paste() -> str:
    import pyautogui; pyautogui.hotkey("ctrl","v"); return "Вставлено."
def clipboard_read()  -> str:
    try:
        import pyperclip; t = pyperclip.paste()
        return f"В буфере: {t[:200]}" if t else "Буфер пуст."
    except Exception: return None

def switch_window()   -> str:
    import pyautogui; pyautogui.hotkey("alt","tab"); return "Переключаю."
def window_minimize() -> str:
    import pyautogui; pyautogui.hotkey("win","d");   return "Окна свёрнуты."
def window_maximize() -> str:
    import pyautogui; pyautogui.hotkey("win","up");  return "Окно развёрнуто."
def window_close()    -> str:
    import pyautogui; pyautogui.hotkey("alt","f4");  return "Окно закрыто."

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
    try: subprocess.Popen(path); return f"{ru_name} открыт."
    except Exception:
        try: subprocess.Popen(path, shell=True); return f"{ru_name} открыт."
        except Exception: return None

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
        try: subprocess.run(["powershell.exe",
            f"Stop-Process -Name {b} -ErrorAction SilentlyContinue"],check=False)
        except Exception: pass
    return "Браузер закрыт."

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
    if os.path.exists(path): os.startfile(path); return f"{folder_name.capitalize()} открыта."
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
        return (f"В {city}: {d['weather'][0]['description']}, "
                f"{round(d['main']['temp'])}°, ощущается {round(d['main']['feels_like'])}°.")
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
    open("список дел.txt","w",encoding="utf-8").close(); return "Список очищен."

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
        try: subprocess.run(["powershell.exe",
            f"Stop-Process -Name {p} -ErrorAction SilentlyContinue"],check=False)
        except Exception: pass
    return "Музыка остановлена."

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

def stop_alarm()         -> str: global alarm_thread; alarm_thread = None; return "Будильник отключён."
def break_reminder_on()  -> str:
    global break_reminder_active
    if break_reminder_active: return "Напоминания уже включены."
    break_reminder_active = True
    def _r():
        while break_reminder_active:
            time.sleep(1800)
            if break_reminder_active: speak("Ты работаешь 30 минут. Перерыв!")
    threading.Thread(target=_r, daemon=True).start()
    return "Напоминания о перерывах включены."

def break_reminder_off() -> str:
    global break_reminder_active; break_reminder_active = False; return "Напоминания выключены."
def shutdown()           -> str: os.system("shutdown /s /t 10"); return "Выключаю через 10 секунд."
def restart()            -> str: os.system("shutdown /r /t 10"); return "Перезагружаю через 10 секунд."
def sleep_pc()           -> str: os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0"); return "Спящий режим."
def cancel_shutdown()    -> str: os.system("shutdown /a"); return "Выключение отменено."

def dictate() -> str:
    speak("Говори — напечатаю."); t = listen_fn()
    if not t: return None
    try:
        import pyautogui, pyperclip; pyperclip.copy(t); pyautogui.hotkey("ctrl","v"); return "Напечатала."
    except Exception: return None

def find_file() -> str:
    speak("Что ищем?"); name = listen_fn()
    if not name: return None
    try:
        r = subprocess.run(["powershell.exe",
            f"Get-ChildItem -Path $env:USERPROFILE -Recurse -Filter '*{name}*' "
            f"-ErrorAction SilentlyContinue | Select-Object -First 3 FullName | Format-Table -HideTableHeaders"],
            capture_output=True, text=True, timeout=10)
        found = [l.strip() for l in r.stdout.strip().splitlines() if l.strip()]
        return f"Нашла: {found[0]}." if found else f"{name} не найден."
    except Exception: return None

def remind_me(minutes: int, text: str = "Напоминание!") -> str:
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

def break_code(): speak("До встречи!"); exit()


def execute_command(cmd: str, query: str = "") -> str | None:
    if cmd.startswith("open_folder:"): return open_folder(cmd.split(":",1)[1])
    if cmd.startswith("open_app:"):    return open_app(cmd.split(":",1)[1])
    if cmd == "ping":   return random.choice(["Да?", "Я здесь.", "Слушаю."])
    if cmd == "mute":
        global is_muted; is_muted = True; stop_speech(); return ""
    if cmd == "unmute":
        is_muted = False; return "Снова слушаю."
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
    if cmd == "calculate":  return calculate(query)
    if cmd == "remind_me":
        secs = SlotExtractor.time_to_seconds(query)
        return remind_me(secs//60 if secs else 5, query)

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
        "find_file":find_file,"open_folder":open_folder,"dictate":dictate,
        "shutdown":shutdown,"restart":restart,"sleep":sleep_pc,
        "cancel_shutdown":cancel_shutdown,"stop_alarm":stop_alarm,
        "break_reminder_on":break_reminder_on,"break_reminder_off":break_reminder_off,
        "break_code":break_code,
    }
    fn = dispatch.get(cmd)
    return fn() if fn else None


# ─── MAIN ────────────────────────────────────────────────────────────────────

def main():
    global listen_fn, is_muted

    os.makedirs("music", exist_ok=True)
    os.makedirs("sounds", exist_ok=True)
    if not os.path.exists("список дел.txt"):
        open("список дел.txt","w",encoding="utf-8").close()

    _generate_cache()
    _init_silero()
    _vol_init()

    if not PICOVOICE_KEY:
        print("  [!] Добавь PICOVOICE_KEY в .env"); exit(1)
    if not os.path.exists(PPN_PATH):
        print(f"  [!] Файл не найден: {PPN_PATH}"); exit(1)

    # Vosk
    vosk_listener = VoskListener(list(LOCAL_COMMANDS.keys()))
    listen_fn = lambda timeout=8: vosk_listener.listen(timeout)

    # Porcupine
    porcupine = pvporcupine.create(
        access_key=PICOVOICE_KEY,
        keyword_paths=[PPN_PATH],
        sensitivities=[0.5]
    )
    recorder = PvRecorder(device_index=-1, frame_length=porcupine.frame_length)
    recorder.start()

    play_cached("ready")
    print("\n  Говори 'Эй Лора'\n")

    def _process(query: str):
        global is_muted
        if is_speaking: stop_speech()
        query = query.lower().strip()
        t0    = time.time()

        if any(w in query for w in STOP_TRIGGERS):
            break_code()

        # Пинг — "лора" в середине диалога
        if "лора" in query or "lora" in query:
            reply = random.choice(["Да?", "Я здесь.", "Слушаю."])
            speak(reply); return

        # Размут — до проверки мута
        if any(w in query for w in UNMUTE_TRIGGERS):
            if is_muted:
                is_muted = False
                reply = random.choice(["Снова слушаю.", "Да?", "Я здесь."])
                speak(reply)
            return

        # Мут
        if any(w in query for w in MUTE_TRIGGERS):
            is_muted = True; stop_speech(); return

        if is_muted: return

        # Точное совпадение (Vosk с грамматикой даёт именно фразу из словаря)
        cmd = LOCAL_COMMANDS.get(query)

        # Частичное совпадение на случай небольших отклонений
        if not cmd:
            for phrase, c in LOCAL_COMMANDS.items():
                if phrase in query or query in phrase:
                    cmd = c; break

        if cmd:
            result = execute_command(cmd, query)
            if result is None:
                play_unclear()
            elif result == "":
                pass   # мут — молчим
            else:
                speak(result)
        else:
            play_unclear()

    try:
        while True:
            pcm = recorder.read()
            if porcupine.process(pcm) >= 0:
                recorder.stop()
                _log_you("эй лора")
                if is_muted:
                    is_muted = False
                play_wake()
                last_active = time.time()

                cmd_text = vosk_listener.listen(timeout=6)
                recorder.start()

                if not cmd_text:
                    continue

                _log_you(cmd_text)
                _process(cmd_text)
                last_active = time.time()

                # Окно активности — следующие команды без wake word
                while (time.time() - last_active) < WINDOW_AFTER_AI:
                    followup = vosk_listener.listen(timeout=2)
                    if followup:
                        last_active = time.time()
                        _log_you(followup)
                        _process(followup)

    except KeyboardInterrupt:
        print("\n  Завершение...")
    finally:
        recorder.stop()
        recorder.delete()
        porcupine.delete()


if __name__ == "__main__":
    main()
