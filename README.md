# AutoAudioBook
<img width="913" height="621" alt="tagging" src="https://github.com/user-attachments/assets/22519db0-0105-4db8-8f27-edfca3f407ed" />

AutoAudioBook ingests PDF or DOCX books, extracts the text, creates a reviewable annotation artifact with Gemini inline TTS tags, previews chunking, and can generate per-chunk WAV audio with Gemini 3.1 Flash TTS.
I run this on a Ubuntu 24.04 VM running on two cores and 4GB ram om a Beelink ME Pro.

This project was developed in VS Code with GitHub Copilot using GPT-5.4.

## Install on Ubuntu 24.04

1. Clone the repository:

```bash
git clone https://github.com/tronba/AutoAudioBook.git
cd AutoAudioBook
```

2. Run the installer:

```bash
sudo bash install_ubuntu_24.sh
```

3. Open the app in a browser:

```text
http://<server-ip>:8000
```

The installer will:

- install Ubuntu packages
- create or reuse `.venv`
- install Python dependencies
- securely prompt for the Gemini API key
- save the key to `/etc/autoaudiobook/autoaudiobook.env` with restricted permissions
- install and enable the `systemd` service

To manage the service:

```bash
sudo systemctl status autoaudiobook
sudo systemctl restart autoaudiobook
sudo journalctl -u autoaudiobook -n 100 --no-pager
```

## What it does
<img width="1672" height="941" alt="model" src="https://github.com/user-attachments/assets/ff0f7a4d-c5df-4d55-a17f-e000f2da5cc7" />

- Imports PDF or DOCX books
- Generates a reviewable annotated DOCX with inline TTS tags
- Lets you upload an approved annotation for audio generation
- Previews chunking before synthesis
- Generates WAV audio with Gemini TTS
![Uploading tagging.PNG…]()
![Uploading 356febc0-2429-417b-8730-b642ade9c7c7.png…]()


## Input file notes

- DOCX input files should mark chapter headings with the Word style `Header 1`
- Text before the first Chapter 1 marker is included in Chapter 1 by default
- That opening text can be split out instead if you enable the separate pre-chapter text option during generation

## Tag configuration

Editable inline tag vocabulary lives in `tts_tags.toml`.

```toml
[[expressive_tags]]
tag = "[angry]"
min_mode = "expressive"
```

Each tag belongs to either `expressive_tags` or `vocalization_tags`, and each entry includes a `min_mode` of `conservative`, `balanced`, or `expressive`.

## Storage layout

- `storage/app.db` - SQLite database
- `storage/uploads/` - original uploaded source files
- `storage/extracted/` - normalized extracted JSON
- `storage/annotated/` - draft and approved annotation DOCX files
- `storage/audio/` - reserved for future chunk and chapter audio output

## Gemini configuration

Set these environment variables on the server:

- `GEMINI_API_KEY`
- `GEMINI_TEXT_MODEL` optional, defaults to `gemini-2.5-flash`
- `GEMINI_TTS_MODEL` optional, defaults to `gemini-3.1-flash-tts-preview`
