"""Transcribe a 16kHz mono WAV using a local sherpa-onnx Whisper model.

Whisper reads a fixed 30-second window, so longer audio is split into
chunks and each chunk is timestamped from its position in the file.
"""
import argparse
import wave
from pathlib import Path

import numpy as np
import sherpa_onnx

SAMPLE_RATE = 16000


def read_wav(path):
    with wave.open(str(path), "rb") as handle:
        if handle.getnchannels() != 1 or handle.getframerate() != SAMPLE_RATE:
            raise SystemExit(
                f"expected mono {SAMPLE_RATE}Hz, got {handle.getnchannels()}ch "
                f"{handle.getframerate()}Hz"
            )
        raw = handle.readframes(handle.getnframes())
    return np.frombuffer(raw, dtype=np.int16).astype(np.float32) / 32768.0


def timestamp(seconds):
    return f"{int(seconds) // 60:02d}:{int(seconds) % 60:02d}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("audio")
    ap.add_argument("--model-dir", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--chunk", type=float, default=28.0,
                    help="seconds per chunk; must stay under Whisper's 30s window")
    args = ap.parse_args()

    model = Path(args.model_dir)
    prefix = model.name.replace("sherpa-onnx-whisper-", "")
    recognizer = sherpa_onnx.OfflineRecognizer.from_whisper(
        encoder=str(model / f"{prefix}-encoder.int8.onnx"),
        decoder=str(model / f"{prefix}-decoder.int8.onnx"),
        tokens=str(model / f"{prefix}-tokens.txt"),
        num_threads=4,
    )

    samples = read_wav(args.audio)
    step = int(args.chunk * SAMPLE_RATE)
    lines = []

    for index, start in enumerate(range(0, len(samples), step)):
        chunk = samples[start:start + step]
        if len(chunk) < SAMPLE_RATE // 2:
            break
        stream = recognizer.create_stream()
        stream.accept_waveform(SAMPLE_RATE, chunk)
        recognizer.decode_stream(stream)
        text = stream.result.text.strip()
        mark = timestamp(start / SAMPLE_RATE)
        print(f"[{mark}] {text}", flush=True)
        if text:
            lines.append(f"[{mark}] {text}")

    Path(args.out).write_text("\n".join(lines) + "\n")
    print(f"\nwrote {args.out} ({len(lines)} segments)")


if __name__ == "__main__":
    main()
