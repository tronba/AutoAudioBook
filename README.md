# AutoAudioBook

Version 0.4 is the current working build. It is a local-first single-page app that ingests PDF or DOCX books, extracts text, uses Gemini to create a reviewable annotation artifact, previews chunking, and can generate per-chunk WAV files with Gemini 3.1 Flash TTS.

## Current scope

- Upload PDF or DOCX files
- Extract text and infer rough chapter boundaries
- Generate a Gemini-annotated DOCX that keeps the original text and inserts inline Gemini TTS tags only
- Re-upload an approved annotation DOCX in that same inline-tag format
- Preview paragraph-aware chunking
- Preview the exact Gemini TTS prompt for a selected chunk
- Generate WAV audio for a selected approved chunk with Gemini 3.1 Flash TTS
- Persist state in SQLite and files on local disk

ffmpeg chapter merging is intentionally not wired yet. The current code now includes a Gemini-first annotation and TTS path against the existing internal schema.

## Run locally

1. Create a virtual environment.
2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Start the app:

```bash
uvicorn app:app --reload
```

4. Open http://127.0.0.1:8000

## Run on Ubuntu 24.04 test server

Use the committed server guide in [docs/ubuntu-24-server-setup.md](docs/ubuntu-24-server-setup.md).

The short version is:

```bash
sudo bash install_ubuntu_24.sh
```

That installer:

- installs Ubuntu packages
- creates or reuses `.venv`
- installs Python dependencies
- securely prompts for the Gemini API key
- saves the key to `/etc/autoaudiobook/autoaudiobook.env` with restricted permissions
- installs and enables the `systemd` service

To check or manage the service later on Ubuntu:

```bash
sudo systemctl status autoaudiobook
sudo systemctl restart autoaudiobook
sudo journalctl -u autoaudiobook -n 100 --no-pager
```

## Tag configuration

Editable inline tag vocabulary now lives in the root config file `tts_tags.toml`.

- `expressive_tags` contains delivery-style tags
- `vocalization_tags` contains sound or vocalization tags
- each tag entry includes a `min_mode` value of `conservative`, `balanced`, or `expressive`

Example:

```toml
[[expressive_tags]]
tag = "[angry]"
min_mode = "expressive"
```

This lets a user amend the accessible tags without changing Python code.

## Storage layout

- `storage/app.db` - SQLite database
- `storage/uploads/` - original uploaded source files
- `storage/extracted/` - normalized extracted JSON
- `storage/annotated/` - draft and approved annotation DOCX files
- `storage/audio/` - reserved for future chunk and chapter audio output

## Next implementation targets

- Merge successful chunk WAV files into chapter outputs
- Add ffmpeg-based chapter MP3 assembly
- Add end-to-end retry and status flow for generation jobs

## Gemini configuration

Set these environment variables on the server:

- `GEMINI_API_KEY`
- `GEMINI_TEXT_MODEL` optional, defaults to `gemini-2.5-flash`
- `GEMINI_TTS_MODEL` optional, defaults to `gemini-3.1-flash-tts-preview`

## Research notes

- Gemini TTS integration notes: [docs/gemini-tts-research.md](docs/gemini-tts-research.md)
- Gemini TTS tag reference: [docs/gemini-tts-tag-reference.md](docs/gemini-tts-tag-reference.md)
- LLM annotation pipeline notes: [docs/llm-annotation-plan.md](docs/llm-annotation-plan.md)
