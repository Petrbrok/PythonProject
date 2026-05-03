# 🎙️ Лора — голосовой ассистент на Python

![Python](https://img.shields.io/badge/Python-3.10+-blue?logo=python&logoColor=white)
![Groq](https://img.shields.io/badge/LLM-Groq%20LLaMA%203.3-orange?logo=meta&logoColor=white)
![TTS](https://img.shields.io/badge/TTS-edge--tts%20Microsoft-blueviolet?logo=microsoft)
![STT](https://img.shields.io/badge/STT-Vosk%20офлайн-green)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?logo=windows)
![License](https://img.shields.io/badge/License-MIT-brightgreen)

Голосовой ассистент **Лора** — реагирует на wake word «Эй Лора», распознаёт команды офлайн через Vosk, отвечает через кешированный TTS без задержки. При неизвестной фразе отвечает через LLaMA 3.3 (Groq).

---

## ✨ Возможности

| Категория | Что умеет |
|---|---|
| 🎤 **Wake word** | Реагирует на «Эй Лора» через Vosk с узкой грамматикой |
| 🧠 **ИИ** | Свободные вопросы через LLaMA 3.3 70B (Groq API) |
| 🔊 **TTS** | edge-tts Microsoft Neural, все ответы кешированы — нет задержки |
| 💻 **Система** | Громкость, яркость, скриншот, WiFi, питание |
| 📱 **Приложения** | Открыть/закрыть Telegram, Discord, Chrome, VS Code и др. |
| 📁 **Папки** | Открыть загрузки, документы, рабочий стол |
| 📋 **Задачи** | Добавить, показать, очистить список дел |
| 🌍 **Погода** | Текущая погода по городу (OpenWeatherMap) |
| 🌙 **Режимы** | Ночной, презентация, утренний — запускают набор действий |
| 🔇 **Мут** | Голосом заглушить и разбудить ассистента |

---

## 🚀 Быстрый старт

### 1. Клонировать репозиторий
```bash
git clone https://github.com/Petrbrok/PythonProject.git
cd PythonProject
```

### 2. Установить зависимости
```bash
pip install -r requirements.txt
```

### 3. Скачать модель Vosk
Скачать [vosk-model-small-ru-0.22](https://alphacephei.com/vosk/models), распаковать и переименовать папку в `model` рядом с `project1.py`.

### 4. Настроить .env
```env
GROQ_API_KEY=твой_ключ        # console.groq.com — бесплатно
WEATHER_API_KEY=твой_ключ     # openweathermap.org — опционально
EDGE_VOICE=ru-RU-SvetlanaNeural
```

### 5. Запустить
```bash
python project1.py
```

Скажи **«Эй Лора»** — ассистент готов к работе.

---

## 🏗️ Архитектура

```
Микрофон
   │
   ├─► WakeWordListener (Vosk, грамматика ["эй лора"])
   │         │ wake word обнаружен
   │         ▼
   └─► VoskListener (свободная речь)
             │
             ├─► Точное совпадение с LOCAL_COMMANDS
             ├─► Rapidfuzz (нечёткое совпадение, порог 80%)
             ├─► Частичное совпадение (3+ слова)
             └─► Groq LLaMA 3.3 (если команда не найдена)
                       │
                       ▼
              TTS кеш (edge-tts mp3) → pygame
```

---

## 🗣️ Примеры команд

```
"Эй Лора, который час"
"Эй Лора, открой телеграм"
"Эй Лора, сделай скриншот"
"Эй Лора, ночной режим"
"Эй Лора, режим презентации"
"Эй Лора, добавь задачу купить продукты"
"Эй Лора, погода в Москве"
"Эй Лора, список команд"       → таблица в консоли
"Эй Лора, замолчи"             → отключение ответов
"размут" / "слушай"            → повторное включение ответов
```

---

## 📁 Структура проекта

```
PythonProject/
├── project1.py          # Основной файл
├── commands.yaml        # Кастомные фразы (не трогая код)
├── requirements.txt     # Зависимости
├── .env                 # API ключи (не коммитить)
├── model/               # Модель Vosk (скачать отдельно)
├── sounds/              # Кеш TTS (генерируется автоматически)
│   ├── cache/           # Системные фразы
│   ├── wake/            # Wake-ответы
│   └── resp/            # Ответы команд
├── music/               # Музыка для воспроизведения
└── список дел.txt       # Создаётся автоматически
```

---

## 🔑 API ключи

| Сервис | Где получить | Бесплатно |
|---|---|---|
| **Groq** | [console.groq.com](https://console.groq.com) | ✅ |
| **OpenWeatherMap** | [openweathermap.org](https://openweathermap.org/api) | ✅ |
| **edge-tts** | Не нужен | ✅ |

---

## 🛠️ Стек

- **[Vosk](https://alphacephei.com/vosk/)** — офлайн STT, русская модель
- **[Groq](https://groq.com/)** — быстрый инференс LLaMA 3.3 70B
- **[edge-tts](https://github.com/rany2/edge-tts)** — Microsoft Neural TTS
- **[rapidfuzz](https://github.com/maxbachmann/RapidFuzz)** — нечёткое сопоставление команд
- **[pygame](https://www.pygame.org/)** — воспроизведение аудио
- **[sounddevice](https://python-sounddevice.readthedocs.io/)** — захват микрофона

---

## 📄 Лицензия

MIT