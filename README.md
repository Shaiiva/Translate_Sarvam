# Audio Translation & TTS Pipeline

A Python pipeline that transcribes English audio, translates it into multiple Indian languages using Sarvam AI, and optionally generates speech audio.

## Features

- Audio → Transcript (Whisper)
- Transcript → Translation (Sarvam AI)
- Translation → Speech (Sarvam AI)
- Resume interrupted projects
- Sequential or Parallel processing
- OGG, MP3, and WAV output
- Supports Hindi, Telugu, Marathi, Punjabi, Tamil, and Bengali

## Requirements

- Python 3.10 or later
- FFmpeg installed and added to PATH
- Sarvam AI API Key

## Installation

Install the required Python packages:

```bash
pip install -r requirements.txt
```

## Run

```bash
python pipeline.py
```

The pipeline will prompt for:

- Sarvam API Key
- Input Type (Audio or Transcript)
- Project Folder
- Processing Mode
- Target Languages
- Audio Output Format

## Project Folder

### Audio Mode

Place your audio files directly inside the project folder.

### Transcript Mode

Place one Excel file inside the project folder with these headers:

- File Name
- Source Transcript

The pipeline automatically creates the required output folders and Excel files.

## Notes

- FFmpeg must be installed and available in your system PATH.
- Use your own Sarvam AI API key.
- Whisper runs on CPU by default.