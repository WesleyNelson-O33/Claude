# Payroll Tax recording

Source: the `payrolltax` release on this repo, five phone recordings of the
Payroll Tax / Payroll Rec session of 25 August 2026, 85 minutes in total.

Only the transcripts are committed. The video, the extracted audio and the
sampled frames are ignored, because the originals already live on the release
and GitHub will not take files that size in the tree.

To rebuild them, download the release assets into `raw/`, then extract 16kHz
mono audio and run `tools/transcribe.py` against the sherpa-onnx Whisper
small.en model.
