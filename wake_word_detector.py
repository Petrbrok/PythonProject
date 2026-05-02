"""
DTW wake word detector — аналог Raven, реализован на Python.
Сравнивает входящий аудио с записанными шаблонами через MFCC + DTW.
"""
import os
import wave
import numpy as np
import threading
import queue
import sounddevice as sd
from python_speech_features import mfcc
from scipy.spatial.distance import cdist

SAMPLE_RATE    = 16000
FRAME_SIZE     = 1024         # ~64мс
WINDOW_SEC     = 2.5          # окно сравнения
WINDOW_FRAMES  = int(SAMPLE_RATE * WINDOW_SEC)
THRESHOLD      = 0.55         # порог срабатывания (0-1), выше = менее чувствительно
MIN_MATCHES    = 2            # сколько шаблонов должно совпасть
TEMPLATES_DIR  = "wake_word_templates"


def _load_wav(path: str) -> np.ndarray:
    with wave.open(path, "r") as wf:
        raw = wf.readframes(wf.getnframes())
    return np.frombuffer(raw, dtype=np.int16).astype(np.float32) / 32768.0


def _compute_mfcc(audio: np.ndarray) -> np.ndarray:
    return mfcc(audio, samplerate=SAMPLE_RATE, numcep=13, nfilt=26, nfft=512)


def _dtw_distance(a: np.ndarray, b: np.ndarray) -> float:
    """DTW расстояние между двумя MFCC последовательностями."""
    n, m = len(a), len(b)
    cost = cdist(a, b, metric="euclidean")
    dtw  = np.full((n + 1, m + 1), np.inf)
    dtw[0, 0] = 0.0
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            dtw[i, j] = cost[i-1, j-1] + min(dtw[i-1, j], dtw[i, j-1], dtw[i-1, j-1])
    return dtw[n, m] / (n + m)


def _score(distance: float) -> float:
    """Переводим расстояние в вероятность 0-1."""
    return max(0.0, 1.0 - distance / 10.0)


class WakeWordDetector:
    """
    Аналог Raven — DTW по MFCC шаблонам.
    Использование:
        detector = WakeWordDetector(callback=my_func)
        detector.start()
        ...
        detector.stop()
    """

    def __init__(self, callback, templates_dir=TEMPLATES_DIR,
                 threshold=THRESHOLD, min_matches=MIN_MATCHES):
        self.callback      = callback
        self.threshold     = threshold
        self.min_matches   = min_matches
        self._running      = False
        self._thread       = None
        self._q            = queue.Queue()

        # Загружаем шаблоны
        self._templates = []
        if not os.path.exists(templates_dir):
            raise RuntimeError(
                f"Папка шаблонов не найдена: {templates_dir}\n"
                f"Запусти record_wake_word.py чтобы записать шаблоны."
            )
        for fname in sorted(os.listdir(templates_dir)):
            if fname.endswith(".wav"):
                path  = os.path.join(templates_dir, fname)
                audio = _load_wav(path)
                feat  = _compute_mfcc(audio)
                self._templates.append(feat)

        if not self._templates:
            raise RuntimeError(
                f"В папке {templates_dir} нет wav файлов.\n"
                f"Запусти record_wake_word.py чтобы записать шаблоны."
            )

        print(f"  [wake] Загружено шаблонов: {len(self._templates)}")

    def _audio_callback(self, indata, frames, time, status):
        self._q.put(indata.copy().flatten())

    def _detect_loop(self):
        buffer = np.zeros(WINDOW_FRAMES, dtype=np.float32)
        refractory = 0   # кадров до следующего срабатывания

        with sd.InputStream(
            samplerate=SAMPLE_RATE,
            channels=1,
            dtype="float32",
            blocksize=FRAME_SIZE,
            callback=self._audio_callback
        ):
            while self._running:
                try:
                    chunk = self._q.get(timeout=0.1)
                except queue.Empty:
                    continue

                # Сдвигаем буфер
                buffer = np.roll(buffer, -len(chunk))
                buffer[-len(chunk):] = chunk

                if refractory > 0:
                    refractory -= 1
                    continue

                # Сравниваем буфер с шаблонами
                feat    = _compute_mfcc(buffer)
                matches = 0
                for tmpl in self._templates:
                    dist  = _dtw_distance(feat, tmpl)
                    score = _score(dist)
                    if score >= self.threshold:
                        matches += 1

                if matches >= self.min_matches:
                    refractory = int(SAMPLE_RATE / FRAME_SIZE * 2)  # 2 сек пауза
                    self.callback()

    def start(self):
        self._running = True
        self._thread  = threading.Thread(target=self._detect_loop, daemon=True)
        self._thread.start()

    def stop(self):
        self._running = False
        if self._thread:
            self._thread.join(timeout=2)
