# 🎙️ Лора — голосовой ассистент на Python

![Python](https://img.shields.io/badge/Python-3.10+-blue?logo=python&logoColor=white)
![Groq](https://img.shields.io/badge/LLM-Groq%20LLaMA%203.3-orange?logo=meta&logoColor=white)
![TTS](https://img.shields.io/badge/TTS-edge--tts%20Microsoft-blueviolet?logo=microsoft)
![STT](https://img.shields.io/badge/STT-Google%20Speech-green?logo=google)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?logo=windows)
![License](https://img.shields.io/badge/License-MIT-brightgreen)

Голосовой ассистент с именем **Лора** — реагирует на своё имя, понимает естественную речь, выполняет команды и отвечает как живой собеседник. Проект полностью на русском языке.

---

## ✨ Возможности

| Категория | Что умеет |
|---|---|
| 🧠 **ИИ** | Ответы через LLaMA 3.3 70B (Groq API), понимает контекст |
| 🎤 **Голос** | Распознавание речи через Google Speech Recognition |
| 🔊 **Речь** | Синтез через Microsoft edge-tts (`SvetlanaNeural`) |
| 💻 **Система** | Открыть/закрыть приложения, скриншот, громкость, сон, перезагрузка |
| 🌐 **Браузер** | Открыть сайт или поиск Google голосом |
| 📋 **Задачи** | Добавить, прочитать, очистить список дел |
| ⏱️ **Таймер** | Установить таймер голосом |
| 🎵 **Музыка** | Воспроизведение из локальной папки |
| 🌍 **Перевод** | Перевод фраз на английский |
| 🔇 **Мут** | Голосом заглушить/разбудить ассистента |

---

## 🚀 Быстрый старт

### 1. Клонировать репозиторий
```bash
git clone https://github.com/ВАШ_НИК/laura-assistant.git
cd laura-assistant
```

### 2. Установить зависимости
```bash
pip install -r requirements.txt
```

### 3. Настроить переменные окружения
```bash
cp .env.example .env
# Открой .env и вставь свои ключи
```

### 4. Запустить
```bash
python project1.py
```

Скажите **«Лора»** — и ассистент готов к работе.

---

## 🔑 Получение API ключей

| Сервис | Где получить | Бесплатный тариф |
|---|---|---|
| **Groq API** | [console.groq.com](https://console.groq.com) | ✅ Да |
| **edge-tts** | Встроен, ключи не нужны | ✅ Бесплатно |

---

## 📁 Структура проекта

```
laura-assistant/
├── project1.py          # Основной файл ассистента
├── requirements.txt     # Зависимости
├── .env.example         # Шаблон переменных окружения
├── .env                 # Твои ключи (не коммитить!)
├── .gitignore
├── music/               # Папка с музыкой (mp3/wav)
├── sounds/              # Звуки подтверждения команд
└── список дел.txt       # Автосоздаётся при запуске
```

---

## 🗣️ Примеры команд

```
"Лора, который час?"
"Лора, открой телеграм"
"Лора, сделай скриншот"
"Лора, поставь таймер на 5 минут"
"Лора, добавь задачу купить продукты"
"Лора, переведи фразу"
"Лора, открой ютуб"
"Лора, выключи компьютер"
"Лора, замолчи"           → мут
"Проснись"                → размут
```

---

## ⚙️ Настройка голоса

В файле `.env` можно изменить голос TTS:

```env
# Женский (по умолчанию)
EDGE_VOICE=ru-RU-SvetlanaNeural

# Мужской
EDGE_VOICE=ru-RU-DmitryNeural
```

Все доступные голоса: `edge-tts --list-voices`

---

## 🛠️ Стек технологий

- **[Groq](https://groq.com/)** — сверхбыстрый инференс LLaMA 3.3 70B
- **[edge-tts](https://github.com/rany2/edge-tts)** — синтез речи Microsoft Neural TTS
- **[SpeechRecognition](https://github.com/Uberi/speech_recognition)** — обёртка над Google STT
- **[pygame](https://www.pygame.org/)** — воспроизведение аудио
- **[python-dotenv](https://github.com/theskumar/python-dotenv)** — управление переменными окружения

---

## 📄 Лицензия

MIT — используй, форкай, улучшай.
