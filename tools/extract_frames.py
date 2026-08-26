"""Extract candidate screenshots from a screen-recording.

Pulls two frame sets:
  scene/  - frames at detected scene changes (where the screen actually changed)
  grid/   - frames at a fixed interval, as a fallback for slow-moving footage
"""
import argparse
import json
import subprocess
import sys
from pathlib import Path

import imageio_ffmpeg

FFMPEG = imageio_ffmpeg.get_ffmpeg_exe()
FFPROBE = FFMPEG.replace("ffmpeg", "ffprobe")


def run(cmd):
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        sys.exit(f"command failed: {' '.join(cmd)}\n{result.stderr[-2000:]}")
    return result.stdout


def probe(video):
    raw = run([FFPROBE, "-v", "error", "-show_streams", "-show_format",
               "-of", "json", str(video)])
    data = json.loads(raw)
    video_stream = next(s for s in data["streams"] if s["codec_type"] == "video")
    return {
        "duration": float(data["format"]["duration"]),
        "width": int(video_stream["width"]),
        "height": int(video_stream["height"]),
        "codec": video_stream["codec_name"],
        "has_audio": any(s["codec_type"] == "audio" for s in data["streams"]),
    }


def extract(video, outdir, vf, prefix):
    outdir.mkdir(parents=True, exist_ok=True)
    run([FFMPEG, "-y", "-i", str(video), "-vf", vf, "-vsync", "vfr",
         "-frame_pts", "1", "-q:v", "2", str(outdir / f"{prefix}_%05d.jpg")])
    return sorted(outdir.glob("*.jpg"))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("video")
    ap.add_argument("--outdir", default="frames")
    ap.add_argument("--scene", type=float, default=0.12,
                    help="scene-change sensitivity, 0-1; lower finds more frames")
    ap.add_argument("--interval", type=float, default=15.0,
                    help="seconds between fallback grid frames")
    ap.add_argument("--width", type=int, default=1280,
                    help="downscale frames to this width")
    args = ap.parse_args()

    video = Path(args.video)
    out = Path(args.outdir)
    info = probe(video)
    print(json.dumps(info, indent=2))

    scale = f"scale={args.width}:-2"
    scene = extract(video, out / "scene",
                    f"select='gt(scene,{args.scene})',{scale}", "scene")
    grid = extract(video, out / "grid",
                   f"fps=1/{args.interval},{scale}", "grid")

    if info["has_audio"]:
        audio = out / "audio.wav"
        run([FFMPEG, "-y", "-i", str(video), "-vn",
             "-ac", "1", "-ar", "16000", str(audio)])
        print(f"audio:  {audio}")
    else:
        print("audio:  none in source")

    print(f"scene frames: {len(scene)}")
    print(f"grid frames:  {len(grid)}")


if __name__ == "__main__":
    main()
