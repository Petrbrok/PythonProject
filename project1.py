import os
import json
import time
import random
import asyncio
import tempfile
import subprocess
import threading
import webbrowser
import speech_recognition
from datetime import datetime
from dotenv import load_dotenv
from groq import Groq
from translate import Translator
import edge_tts
import pygame

load_dotenv()

ai_client = Groq(api_key=os.getenv("GROQ_API_KEY"))

# Голос edge-tts — меняй по вкусу:
# ru-RU-SvetlanaNeural  — женский, мягкий (рекомендуется для Лоры)
# ru-RU-DmitryNeural    — мужской
EDGE_VOICE = os.getenv("EDGE_VOICE", "ru-RU-SvetlanaNeural")

pygame.mixer.init()

sr = speech_recognition.Recognizer()
sr.pause_threshold = 0.8

is_muted = False
is_speaking = False
speak_lock = threading.Lock()
stop_speaking_event = threading.Event()

MUTE_TRIGGERS = (
    "мут", "mute", "mut", "замолчи", "молчи",
    "тихо", "пауза", "не слушай", "заткнись", "стоп"
)

UNMUTE_TRIGGERS = (
    "размут", "unmute", "разм", "размьют", "включись",
    "разуму", "слушай", "продолжай", "ты здесь", "проснись",
    "отмут", "вернись", "активируйся"
)

STOP_TRIGGERS = (
    "завершить работу", "заверши работу", "выключись",
    "завершись", "закройся", "выход", "exit", "quit", "пока",
    "до свидания", "отключись", "выключи себя"
)

NAME_TRIGGERS = (
    "лора", "флора", "laura", "лёра", "лаура", "лор"
)

# Пути к приложениям — добавь свои если нужно
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
    "task manager":    "taskmgr.exe",
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

SYSTEM_PROMPT = """Тебя зовут Лора — голосовой ассистент с дружелюбным характером. Общаешься как друг — просто, с лёгким юмором, без официоза. Отвечай на русском языке. Максимум 1-2 предложения. Только прямой ответ, без лишних объяснений.

Если тебя спрашивают на чём ты работаешь или кто тебя создал — отвечай что ты Лора, голосовой ассистент, и не раскрывай технические детали.

Не упоминай время в ответе если тебя об этом не спрашивали.
Когда называешь время — говори просто "Сейчас [время]", без обращений и без упоминания UTC.

Если запрос соответствует команде — верни ТОЛЬКО JSON без лишнего текста:
{"command": "название_команды"}

Доступные команды:
- create_task — добавить задачу в список дел
- show_tasks — прочитать список задач
- clear_tasks — очистить список задач
- play_music — включить музыку
- stop_music — остановить музыку
- sound_off — выключить звук системы
- sound_on — включить звук системы
- volume_up — увеличить громкость
- volume_down — уменьшить громкость
- translate — перевести фразу на английский
- open_browser — открыть сайт в браузере
- close_browser — закрыть браузер
- show_help — показать список команд
- get_time — узнать текущее время (только если не указан город или страна)
- get_date — узнать текущую дату
- screenshot — сделать скриншот
- clipboard_read — прочитать текст из буфера обмена
- shutdown — выключить компьютер
- restart — перезагрузить компьютер
- sleep — перевести компьютер в сон
- cancel_shutdown — отменить выключение
- break_code — завершить работу полностью
- window_minimize — свернуть текущее окно
- window_maximize — развернуть текущее окно
- window_close — закрыть текущее окно
- clipboard_copy — скопировать выделенный текст
- clipboard_paste — вставить текст из буфера
- clipboard_write — записать текст в буфер голосом
- brightness_up — увеличить яркость экрана
- brightness_down — уменьшить яркость экрана
- get_weather — узнать погоду
- find_file — найти файл на компьютере
- open_folder — открыть папку по названию
- dictate — продиктовать текст (вставить в активное поле)
- remind_me — напомнить через X минут
- switch_window — переключить окно (Alt+Tab): "следующее окно", "переключи окно", "смени окно"
- window_switch — то же что switch_window
- clipboard_copy — скопировать выделенный текст
- clipboard_paste — вставить из буфера
- clipboard_write — записать текст в буфер голосом
- window_minimize — свернуть текущее окно
- window_maximize — развернуть текущее окно
- window_close — закрыть текущее окно
- get_battery — узнать заряд батареи
- get_cpu — узнать загрузку процессора
- wifi_off — отключить wifi
- wifi_on — включить wifi
- brightness_up — увеличить яркость
- brightness_down — уменьшить яркость
- get_weather — узнать погоду
- dictate — продиктовать текст голосом в активное поле
- find_file — найти файл на компьютере
- open_folder — открыть папку
- notify — показать уведомление на рабочем столе

Если команда open_app — верни JSON с названием:
{"command": "open_app", "app": "название"}

Если команда close_app — верни JSON с названием:
{"command": "close_app", "app": "название"}

Если команда set_timer — верни JSON с секундами:
{"command": "set_timer", "seconds": число}

Если команда remind_me — верни JSON с минутами и текстом:
{"command": "remind_me", "minutes": число, "text": "текст напоминания"}

Если команда open_folder — верни JSON с названием папки:
{"command": "open_folder", "folder": "название"}

Если команда find_file — верни JSON с названием файла:
{"command": "find_file", "name": "название"}

Если спрашивают время в городе или стране — посчитай сам и ответь текстом без JSON. Для страны используй столицу. Для США без города — Вашингтон. Просто называй цифры без упоминания UTC и часовых поясов.
Если запрос не команда — ответь текстом без JSON."""

for folder in ("music", "sounds"):
    if not os.path.exists(folder):
        os.mkdir(folder)
if not os.path.exists("список дел.txt"):
    open("список дел.txt", "w", encoding="utf-8").close()


# ─────────────────────────── РЕЧЬ ───────────────────────────

def _synthesize(text: str):
    """Синтезирует речь через edge-tts, возвращает путь к mp3 или None."""
    async def _run():
        tts = edge_tts.Communicate(text, EDGE_VOICE)
        with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as f:
            tmp = f.name
        await tts.save(tmp)
        return tmp
    try:
        return asyncio.run(_run())
    except Exception as e:
        print(f"  [!] Ошибка синтеза: {e}")
        return None


def speak(text: str):
    """Озвучка с возможностью прерывания через stop_speaking_event."""
    global is_speaking
    if not text:
        return
    print(f"  >> {text}")

    def _speak_worker():
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
        except Exception as e:
            print(f"  [!] Ошибка речи: {e}")
        finally:
            is_speaking = False
            if tmp:
                try:
                    os.unlink(tmp)
                except Exception:
                    pass

    t = threading.Thread(target=_speak_worker, daemon=True)
    t.start()
    t.join()


def speak_async(text):
    """Озвучка без блокировки основного потока."""
    threading.Thread(target=speak, args=(text,), daemon=True).start()


def stop_speech():
    """Немедленно прерывает озвучку."""
    global is_speaking
    stop_speaking_event.set()
    pygame.mixer.music.stop()
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


# ─────────────────────────── МИК ───────────────────────────

def listen(timeout=None):
    """Слушает микрофон. Если is_speaking — прерывает речь при обнаружении голоса."""
    print("  [mic] Слушаю...")
    while True:
        try:
            with speech_recognition.Microphone() as source:
                # Если ассистент говорит — ставим короткий таймаут чтобы реагировать быстрее
                phrase_limit = 8 if not is_speaking else 4
                audio = sr.listen(source, phrase_time_limit=phrase_limit, timeout=timeout)

            # Если что-то услышали пока говорили — сразу прерываем
            if is_speaking:
                stop_speech()

            query = sr.recognize_google(audio, language="ru-RU").lower().strip()
            print(f"  [you] {query}")
            return query
        except speech_recognition.WaitTimeoutError:
            return None
        except speech_recognition.UnknownValueError:
            if is_speaking:
                pass  # Фоновый шум пока говорим — игнорируем
        except Exception:
            pass


# ─────────────────────────── ИИ ───────────────────────────

def ask_ai(query):
    now_local = datetime.now()
    utc_offset = now_local.astimezone().utcoffset()
    offset_hours = int(utc_offset.total_seconds() // 3600)
    now_str = now_local.strftime("%H:%M, %d.%m.%Y")
    try:
        r = ai_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            max_tokens=150,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": f"[системная информация: {now_str} UTC+{offset_hours}] {query}"},
            ],
        )
        return r.choices[0].message.content.strip()
    except Exception as e:
        return f"Ошибка ИИ: {e}"


# ─────────────────────────── КОМАНДЫ ───────────────────────────

def show_help():
    return ("Задачи, музыка, звук, перевод, браузер, скриншот, "
            "время, дата, таймер, буфер обмена, открыть или закрыть приложение, "
            "мут, выключить или перезагрузить компьютер.")

def get_time():
    return f"Сейчас {datetime.now().strftime('%H:%M')}."

def get_date():
    months = ["января","февраля","марта","апреля","мая","июня",
              "июля","августа","сентября","октября","ноября","декабря"]
    now = datetime.now()
    return f"Сегодня {now.day} {months[now.month - 1]} {now.year} года."

def screenshot():
    try:
        import pyautogui
        fname = f"screenshot_{int(time.time())}.png"
        pyautogui.screenshot(fname)
        return f"Скриншот сохранён: {fname}"
    except ImportError:
        return "Установите pyautogui: pip install pyautogui"
    except Exception as e:
        return f"Ошибка: {e}"

def clipboard_read():
    try:
        import pyperclip
        text = pyperclip.paste()
        if text:
            return f"В буфере: {text[:200]}"
        return "Буфер обмена пуст"
    except ImportError:
        return "Установите pyperclip: pip install pyperclip"

def _resolve_app_path(app_name: str) -> str:
    """Возвращает путь к приложению, раскрывая %USERNAME%."""
    raw = APP_PATHS.get(app_name.lower(), app_name)
    username = os.environ.get("USERNAME", os.environ.get("USER", ""))
    return raw.replace("%USERNAME%", username)

def open_app(app_name: str):
    path = _resolve_app_path(app_name)
    try:
        subprocess.Popen(path)
        return f"Открываю {app_name}"
    except FileNotFoundError:
        # Пробуем через shell (работает для встроенных команд)
        try:
            subprocess.Popen(path, shell=True)
            return f"Открываю {app_name}"
        except Exception:
            return f"Не удалось найти {app_name}. Добавьте путь в APP_PATHS."

def close_app(app_name: str):
    # Маппинг имён в процессы
    process_map = {
        "telegram": "telegram", "телеграм": "telegram",
        "discord": "discord", "дискорд": "discord",
        "spotify": "spotify", "спотифай": "spotify",
        "chrome": "chrome", "хром": "chrome",
        "notepad": "notepad", "блокнот": "notepad",
        "word": "winword", "excel": "excel",
        "powerpoint": "powerpnt",
        "obs": "obs64",
        "steam": "steam",
        "vscode": "code", "код": "code",
    }
    proc = process_map.get(app_name.lower(), app_name)
    try:
        subprocess.run(
            ["powershell.exe", f"Stop-Process -Name {proc} -ErrorAction SilentlyContinue"],
            check=False
        )
        return f"Закрываю {app_name}"
    except Exception as e:
        return f"Ошибка: {e}"

def open_browser():
    speak("Какой сайт открыть?")
    site = listen()
    if not site:
        return "Не распознано"
    if "." in site and " " not in site:
        url = f"https://{site}" if not site.startswith("http") else site
    else:
        url = f"https://www.google.com/search?q={site.replace(' ', '+')}"
    webbrowser.open(url)
    return f"Открываю {site}"

def close_browser():
    for browser in ("chrome", "firefox", "msedge", "opera", "brave"):
        try:
            subprocess.run(
                ["powershell.exe", f"Stop-Process -Name {browser} -ErrorAction SilentlyContinue"],
                check=False
            )
        except Exception:
            pass
    return "Браузер закрыт"

def sound_off():
    os.system("powershell.exe (new-object -com wscript.shell).SendKeys([char]173)")
    return "Звук отключён"

def sound_on():
    os.system("powershell.exe (new-object -com wscript.shell).SendKeys([char]173)")
    return "Звук включён"

def volume_up():
    for _ in range(5):
        os.system("powershell.exe (new-object -com wscript.shell).SendKeys([char]175)")
    return "Громкость увеличена"

def volume_down():
    for _ in range(5):
        os.system("powershell.exe (new-object -com wscript.shell).SendKeys([char]174)")
    return "Громкость уменьшена"

def create_task():
    speak("Что добавить в список дел?")
    task = listen()
    if not task:
        return "Не удалось распознать задачу"
    with open("список дел.txt", "a", encoding="utf-8") as f:
        f.write(f"✅ {task}\n")
    return f"Задача добавлена: {task}"

def show_tasks():
    with open("список дел.txt", "r", encoding="utf-8") as f:
        tasks = f.read().strip()
    if not tasks:
        return "Список дел пуст"
    lines = tasks.splitlines()
    return f"В списке {len(lines)} задач: " + "; ".join(l.replace("✅ ", "") for l in lines[:5])

def clear_tasks():
    open("список дел.txt", "w", encoding="utf-8").close()
    return "Список дел очищен"

def play_music():
    files = [f for f in os.listdir("music") if f.endswith((".mp3", ".wav", ".flac"))]
    if not files:
        return "В папке music нет аудиофайлов"
    f = os.path.join("music", random.choice(files))
    os.startfile(f)
    return f"Включаю {os.path.splitext(os.path.basename(f))[0]}"

def stop_music():
    for player in ("wmplayer", "vlc", "spotify", "groove", "musicbee"):
        try:
            subprocess.run(
                ["powershell.exe", f"Stop-Process -Name {player} -ErrorAction SilentlyContinue"],
                check=False
            )
        except Exception:
            pass
    return "Музыка остановлена"

def translate_phrase():
    speak("Скажите фразу для перевода")
    text = listen()
    if not text:
        return "Не распознано"
    try:
        t = Translator(from_lang="ru", to_lang="en")
        return f"По-английски: {t.translate(text)}"
    except Exception as e:
        return f"Ошибка перевода: {e}"

def set_timer(seconds: int):
    def _timer():
        time.sleep(seconds)
        speak("Таймер сработал!")
        play_sound("confirm")
    threading.Thread(target=_timer, daemon=True).start()
    mins, secs = divmod(seconds, 60)
    if mins:
        return f"Таймер на {mins} мин {secs} сек запущен"
    return f"Таймер на {secs} секунд запущен"

def mute():
    global is_muted
    is_muted = True
    stop_speech()
    print("  [muted] Говорите 'размут', 'включись' или 'проснись' чтобы активировать")

def shutdown():
    os.system("shutdown /s /t 10")
    return "Выключаю компьютер через 10 секунд"

def restart():
    os.system("shutdown /r /t 10")
    return "Перезагружаю через 10 секунд"

def sleep_pc():
    os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
    return "Перевожу в сон"

def cancel_shutdown():
    os.system("shutdown /a")
    return "Выключение отменено"

def break_code():
    speak("До встречи!")
    exit()

# ─────────────────────────── НОВЫЕ ФУНКЦИИ ───────────────────────────

def switch_window():
    import pyautogui
    pyautogui.hotkey("alt", "tab")
    return None  # молча

def window_minimize():
    import pyautogui
    pyautogui.hotkey("win", "d")
    return "Сворачиваю"

def window_maximize():
    import pyautogui
    pyautogui.hotkey("win", "up")
    return "Разворачиваю"

def window_close():
    import pyautogui
    pyautogui.hotkey("alt", "f4")
    return "Закрываю окно"

def clipboard_copy():
    import pyautogui
    pyautogui.hotkey("ctrl", "c")
    return "Скопировано"

def clipboard_paste():
    import pyautogui
    pyautogui.hotkey("ctrl", "v")
    return "Вставлено"

def clipboard_write():
    speak("Что записать в буфер?")
    text = listen()
    if not text:
        return "Не распознано"
    try:
        import pyperclip
        pyperclip.copy(text)
        return f"Записал в буфер: {text}"
    except ImportError:
        return "Установите pyperclip"

def get_battery():
    try:
        import psutil
        b = psutil.sensors_battery()
        if b is None:
            return "Батарея не найдена — возможно стационарный ПК"
        status = "заряжается" if b.power_plugged else "на батарее"
        return f"Заряд {int(b.percent)}%, {status}"
    except ImportError:
        return "Установите psutil: pip install psutil"

def get_cpu():
    try:
        import psutil
        cpu = psutil.cpu_percent(interval=1)
        ram = psutil.virtual_memory().percent
        return f"Процессор {cpu}%, оперативная память {ram}%"
    except ImportError:
        return "Установите psutil: pip install psutil"

def wifi_off():
    os.system('netsh interface set interface "Wi-Fi" disabled')
    return "Wi-Fi отключён"

def wifi_on():
    os.system('netsh interface set interface "Wi-Fi" enabled')
    return "Wi-Fi включён"

def brightness_up():
    try:
        import screen_brightness_control as sbc
        cur = sbc.get_brightness()[0]
        sbc.set_brightness(min(100, cur + 20))
        return f"Яркость {min(100, cur + 20)}%"
    except ImportError:
        # Fallback через PowerShell
        os.system("powershell (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,[math]::Min(100,((Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightness).CurrentBrightness)+20))")
        return "Яркость увеличена"

def brightness_down():
    try:
        import screen_brightness_control as sbc
        cur = sbc.get_brightness()[0]
        sbc.set_brightness(max(0, cur - 20))
        return f"Яркость {max(0, cur - 20)}%"
    except ImportError:
        os.system("powershell (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,[math]::Max(0,((Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightness).CurrentBrightness)-20))")
        return "Яркость уменьшена"

def get_weather():
    speak("Какой город?")
    city = listen()
    if not city:
        return "Не распознано"
    api_key = os.getenv("WEATHER_API_KEY")
    if not api_key:
        return "Добавь WEATHER_API_KEY в .env"
    try:
        import urllib.request
        import urllib.parse
        url = (
            f"https://api.openweathermap.org/data/2.5/weather"
            f"?q={urllib.parse.quote(city)}&appid={api_key}&units=metric&lang=ru"
        )
        with urllib.request.urlopen(url, timeout=5) as r:
            data = json.loads(r.read().decode("utf-8"))
        desc    = data["weather"][0]["description"]
        temp    = round(data["main"]["temp"])
        feels   = round(data["main"]["feels_like"])
        humid   = data["main"]["humidity"]
        wind    = round(data["wind"]["speed"])
        return (
            f"В городе {city}: {desc}, {temp}°, ощущается как {feels}°, "
            f"влажность {humid}%, ветер {wind} м/с"
        )
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return f"Город '{city}' не найден"
        return f"Ошибка погоды: {e}"
    except Exception as e:
        return f"Ошибка погоды: {e}"

def dictate():
    speak("Говорите — напечатаю")
    text = listen()
    if not text:
        return "Не распознано"
    try:
        import pyautogui
        import pyperclip
        pyperclip.copy(text)
        pyautogui.hotkey("ctrl", "v")
        return None  # молча вставляем
    except Exception as e:
        return f"Ошибка: {e}"

def find_file():
    speak("Какой файл ищем?")
    name = listen()
    if not name:
        return "Не распознано"
    try:
        result = subprocess.run(
            ["powershell.exe", f"Get-ChildItem -Path $env:USERPROFILE -Recurse -Filter '*{name}*' -ErrorAction SilentlyContinue | Select-Object -First 3 FullName | Format-Table -HideTableHeaders"],
            capture_output=True, text=True, timeout=10
        )
        found = result.stdout.strip()
        if found:
            lines = [l.strip() for l in found.splitlines() if l.strip()]
            return f"Нашёл {len(lines)} файл(ов): {lines[0]}"
        return f"Файл '{name}' не найден"
    except Exception as e:
        return f"Ошибка поиска: {e}"

def open_folder(folder_name: str = ""):
    folders = {
        "документы": os.path.expanduser("~/Documents"),
        "загрузки":  os.path.expanduser("~/Downloads"),
        "рабочий стол": os.path.expanduser("~/Desktop"),
        "десктоп":   os.path.expanduser("~/Desktop"),
        "музыка":    os.path.expanduser("~/Music"),
        "изображения": os.path.expanduser("~/Pictures"),
        "видео":     os.path.expanduser("~/Videos"),
    }
    path = folders.get(folder_name.lower(), folder_name)
    if os.path.exists(path):
        os.startfile(path)
        return f"Открываю {folder_name}"
    return f"Папка '{folder_name}' не найдена"

def notify(text: str = "Уведомление от Лоры"):
    try:
        subprocess.Popen([
            "powershell.exe", "-Command",
            f"[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime] | Out-Null; "
            f"$template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText01); "
            f"$template.GetElementsByTagName('text')[0].AppendChild($template.CreateTextNode('{text}')) | Out-Null; "
            f"$notif = [Windows.UI.Notifications.ToastNotification]::new($template); "
            f"[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('Лора').Show($notif)"
        ])
        return None
    except Exception:
        return None

def remind_me(minutes: int, text: str = "Напоминание!"):
    def _remind():
        time.sleep(minutes * 60)
        speak(f"Напоминаю: {text}")
        notify(text)
    threading.Thread(target=_remind, daemon=True).start()
    return f"Напомню через {minutes} мин: {text}"


COMMANDS = {
    "create_task":    create_task,
    "show_tasks":     show_tasks,
    "clear_tasks":    clear_tasks,
    "play_music":     play_music,
    "stop_music":     stop_music,
    "sound_off":      sound_off,
    "sound_on":       sound_on,
    "volume_up":      volume_up,
    "volume_down":    volume_down,
    "translate":      translate_phrase,
    "open_browser":   open_browser,
    "close_browser":  close_browser,
    "show_help":      show_help,
    "get_time":       get_time,
    "get_date":       get_date,
    "screenshot":     screenshot,
    "clipboard_read": clipboard_read,
    "shutdown":       shutdown,
    "restart":        restart,
    "sleep":          sleep_pc,
    "cancel_shutdown":  cancel_shutdown,
    "break_code":       break_code,
    "switch_window":    switch_window,
    "window_switch":    switch_window,
    "window_minimize":  window_minimize,
    "window_maximize":  window_maximize,
    "window_close":     window_close,
    "clipboard_copy":   clipboard_copy,
    "clipboard_paste":  clipboard_paste,
    "clipboard_write":  clipboard_write,
    "get_battery":      get_battery,
    "get_cpu":          get_cpu,
    "wifi_off":         wifi_off,
    "wifi_on":          wifi_on,
    "brightness_up":    brightness_up,
    "brightness_down":  brightness_down,
    "get_weather":      get_weather,
    "dictate":          dictate,
    "find_file":        find_file,
}


# ─────────────────────────── MAIN ───────────────────────────

def main():
    global is_muted

    print("\n  АССИСТЕНТ ЗАПУЩЕН\n")
    print("  Калибровка микрофона...")
    with speech_recognition.Microphone() as source:
        sr.adjust_for_ambient_noise(source, duration=1.5)
    print("  Готово. Говорите.\n")

    speak("Готова к работе")

    while True:
        query = listen()
        if not query:
            continue

        # Прерываем речь если пользователь заговорил
        if is_speaking:
            stop_speech()
            time.sleep(0.1)

        # Проверка мута
        if is_muted:
            if any(w in query for w in UNMUTE_TRIGGERS):
                is_muted = False
                speak("Снова слушаю!")
            continue

        # Активация по имени
        if not any(name in query for name in NAME_TRIGGERS):
            continue

        # Убираем имя из запроса
        for name in NAME_TRIGGERS:
            query = query.replace(name, "", 1).strip()
        query = query.strip(",. ")
        if not query:
            speak("Слушаю?")
            continue

        # Стоп-триггеры
        if any(w in query for w in STOP_TRIGGERS):
            break_code()
            continue

        # Мут
        if any(w in query for w in MUTE_TRIGGERS):
            mute()
            continue

        # Запрос к ИИ
        response = ask_ai(query)

        try:
            data = json.loads(response)
            command = data.get("command")
            play_sound("confirm")

            if command == "open_app":
                app = data.get("app", "")
                result = open_app(app) if app else "Не указано приложение"
            elif command == "close_app":
                app = data.get("app", "")
                result = close_app(app) if app else "Не указано приложение"
            elif command == "set_timer":
                seconds = int(data.get("seconds", 60))
                result = set_timer(seconds)
            elif command == "remind_me":
                minutes = int(data.get("minutes", 5))
                text = data.get("text", "Напоминание!")
                result = remind_me(minutes, text)
            elif command == "open_folder":
                folder = data.get("folder", "")
                result = open_folder(folder)
            elif command == "find_file":
                name = data.get("name", "")
                if name:
                    speak(f"Ищу {name}")
                    result = find_file.__wrapped__(name) if hasattr(find_file, '__wrapped__') else find_file()
                else:
                    result = find_file()
            elif command and command in COMMANDS:
                result = COMMANDS[command]()
            else:
                result = "Команда не найдена"

            if result:
                speak(result)

        except (json.JSONDecodeError, ValueError):
            speak(response)


if __name__ == "__main__":
    main()