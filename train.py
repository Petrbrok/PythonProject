"""
train.py — запись образцов голоса для улучшения распознавания
Запусти один раз: python train.py
"""

import os
import json
import queue
import time
import vosk
import sounddevice as sd

SAMPLES_FILE = "voice_samples.json"
SAMPLERATE   = 16000
WORDS        = {
    "лора":    ["лора", "флора", "лаура", "лара", "клара", "хлора", "лор"],
    "мут":     ["замолчи", "молчи", "тихо", "стоп", "хватит", "заткнись", "умолкни", "подожди"],
    "размут":  ["слушай", "проснись", "вернись", "включись", "размут", "активируйся"],
}
REPEATS = 4  # сколько раз говорить каждое слово


def load_vosk():
    model_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "model")
    if not os.path.exists(model_path):
        print("❌ Папка model/ не найдена!")
        exit(1)
    vosk.SetLogLevel(-1)
    return vosk.Model(model_path)


def record_sample(model, prompt: str, timeout: int = 4) -> str | None:
    """Записывает одну фразу и возвращает распознанный текст."""
    rec = vosk.KaldiRecognizer(model, SAMPLERATE)
    q   = queue.Queue()

    def _cb(indata, frames, time_info, status):
        q.put(bytes(indata))

    print(f"  🎤 {prompt} (говори сейчас...)")
    result_text = ""

    with sd.RawInputStream(samplerate=SAMPLERATE, blocksize=4000,
                           dtype="int16", channels=1, callback=_cb):
        start = time.time()
        while time.time() - start < timeout:
            try:
                data = q.get(timeout=0.1)
            except queue.Empty:
                continue
            if rec.AcceptWaveform(data):
                r = json.loads(rec.Result())
                result_text = r.get("text", "").strip()
                if result_text:
                    break

        if not result_text:
            r = json.loads(rec.FinalResult())
            result_text = r.get("text", "").strip()

    return result_text if result_text else None


def main():
    print("\n╔══════════════════════════════════════╗")
    print("║   ОБУЧЕНИЕ ЛОРЫ — ЗАПИСЬ ОБРАЗЦОВ    ║")
    print("╚══════════════════════════════════════╝\n")
    print("Этот скрипт записывает как Vosk слышит твой голос")
    print("и сохраняет все варианты для точного распознавания.\n")

    model = load_vosk()
    print("✅ Модель загружена\n")

    # Загружаем существующие образцы если есть
    samples = {}
    if os.path.exists(SAMPLES_FILE):
        with open(SAMPLES_FILE, encoding="utf-8") as f:
            samples = json.load(f)
        print(f"📂 Найден существующий файл с образцами\n")

    for word, default_variants in WORDS.items():
        print(f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
        print(f"  Слово: «{word}»")
        print(f"  Скажи это слово {REPEATS} раза")
        print(f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")

        if word not in samples:
            samples[word] = list(default_variants)

        collected = []
        for i in range(1, REPEATS + 1):
            input(f"\n  Попытка {i}/{REPEATS} — нажми Enter и говори...")
            result = record_sample(model, f"Говори «{word}»")
            if result:
                print(f"  ✅ Распознано: «{result}»")
                if result not in samples[word]:
                    samples[word].append(result)
                    collected.append(result)
                else:
                    print(f"  ℹ️  Уже есть в базе")
            else:
                print(f"  ❌ Не распознано, пропускаем")

        print(f"\n  Добавлено новых вариантов: {len(collected)}")
        print(f"  Всего вариантов для «{word}»: {len(samples[word])}")
        print(f"  Варианты: {', '.join(samples[word])}")

    # Сохраняем
    with open(SAMPLES_FILE, "w", encoding="utf-8") as f:
        json.dump(samples, f, ensure_ascii=False, indent=2)

    print(f"\n✅ Сохранено в {SAMPLES_FILE}")
    print("\nТеперь перезапусти Лору — она автоматически загрузит новые варианты.")


if __name__ == "__main__":
    main()
