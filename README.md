# Audio-to-Text Transcription and Excel Standardization

Transcribe audio recordings with **Groq's Whisper speech-to-text API**, or merge existing Excel transcription
workbooks without making any API calls.

The audio workflow exports its existing three-column format:

| Slide Name | Audio version | Audio to text |
|---|---|---|
| 1.1 Start | 1.1 Start | *(real transcription of A1)* |
| 1.5 Thank You | 1.5 Thank You | *(real transcription of A1)* |
| 1.5 Thank You | 1.5 Thank You | *(real transcription of A2)* |
| 2.1 Service Reception | 2.1 Service Reception | *(real transcription of A1)* |

*Slide Name* and *Audio version* are the filename without its extension and without the trailing `- A1` / `-A2` /
`- a3` suffix (`2.12 Customer Arrival 5 - A2.mp3` becomes `2.12 Customer Arrival 5`). There is one row per audio file;
rows are grouped by slide, slides follow slide-number order (so `2.2` comes before `2.12`), and each slide's audio
files are ordered A1, A2, A3.

The separate **Merge Existing Excel Files** workflow accepts multiple `.xlsx` and `.xls` workbooks. It reads
worksheets containing `Slide Name`, `Audio version`, and `Audio to text` headers (header spacing and capitalization
may vary), preserves workbook/sheet/record order, and downloads exactly four columns: `S.No`, `Slide Name`,
`Audio version`, and `Audio to text`. Serial numbers are generated sequentially in the export; an input `S.No`
column is neither needed nor read. Blank slide names inherit the preceding slide only within the same worksheet.
Possible repeated slide/audio-version pairs are reported for review and are not removed.

## Features

- Multiple files per run; each file is transcribed separately and keeps its own filename.
- Automatic language detection, or pick a language yourself. Transcripts are never translated.
- Progress bar and current-file status; one failed file never stops the others.
- Failed files are listed with the reason and are **not** written to Excel (no fake text, no extra columns).
- Retries for temporary problems (rate limits, timeouts, server errors) with a maximum retry count.
- Excel sheet `Transcriptions`: wrapped text, frozen header row, filters, readable column widths.
- Merge multiple existing Excel transcription workbooks without calling Groq; invalid sheets and possible duplicates
  are reported, and all valid source records are retained.
- Each user supplies a key through the masked application field; keys are isolated to their Streamlit session.

## Prerequisites

- Windows 10/11, [Python 3.10+](https://www.python.org/downloads/) (tick **Add Python to PATH** when installing)
- [VS Code](https://code.visualstudio.com/) with the Python extension (recommended)
- A free Groq account and API key

## Setup (PowerShell)

Open the project folder in VS Code, then open a terminal (**Terminal > New Terminal**) and run:

```powershell
cd path\to\audio-transcription-project
python -m venv venv
.\venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

If activation fails with "running scripts is disabled", run this once and try again:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

In VS Code, press `Ctrl+Shift+P`, choose **Python: Select Interpreter**, and pick the one inside `venv`.

## Get a Groq API key

1. Go to <https://console.groq.com> and sign up or log in.
2. Open **API Keys** (<https://console.groq.com/keys>) and click **Create API Key**.
3. Copy the key (it starts with `gsk_`). It is shown only once.

## Optional local settings

```powershell
Copy-Item .env.example .env
notepad .env
```

No Groq API key belongs in `.env`. Each user enters their own key in the application after it starts. `.env` is for
optional non-secret settings only and is ignored by Git.

Optional settings in `.env`:

| Variable | Default | Meaning |
|---|---|---|
| `GROQ_MODEL` | `whisper-large-v3` | Use `whisper-large-v3-turbo` for a faster, cheaper, slightly less accurate model |
| `MAX_FILE_SIZE_MB` | `25` | Set to `100` only if your Groq account is on the Developer plan |

## Run the application

```powershell
.\venv\Scripts\Activate.ps1
streamlit run app.py
```

The app opens at <http://localhost:8501>. After that, everything happens in the browser.

## Use it

1. On first use, enter your Groq API key and select **Validate and Save**. A successful validation saves it for the
  current session, briefly confirms success, and returns to the main screen. Use **Change API Key** to replace it.
2. Upload one or more audio recordings and optionally choose the audio language; otherwise leave **Auto-detect**.
3. Click **Transcribe and Generate Excel**. The button is disabled while processing, so double clicks are harmless.
4. Review the summary, any failed files, and the preview table.
5. Click **Download Excel File**. A copy is also saved in the `output/` folder.

To standardize existing transcription files, upload multiple Excel workbooks in **Merge Existing Excel Files**,
click **Merge and Standardize**, review the preview and any duplicate/worksheet notices, then download the merged
workbook. This workflow does not require a Groq API key.

## Run the tests

Unit tests use mocked API responses and do not use any Groq quota:

```powershell
pytest -v
```

Real API integration test (uses a little quota; skipped by default). Use a test key in your local shell only:

```powershell
$env:RUN_INTEGRATION = "1"
$env:INTEGRATION_GROQ_API_KEY = "your-test-key"
$env:INTEGRATION_AUDIO = "C:\path\to\short_sample.mp3"
pytest -m integration -v
```

## Deployment

For Streamlit Community Cloud:

1. Push the project source to a GitHub repository. Keep `.env` out of the repository; `.gitignore` excludes it.
2. In [Streamlit Community Cloud](https://share.streamlit.io/), choose **Create app**, select the repository and
  branch, and set the app file path to `app.py`.
3. Deploy with the dependencies from `requirements.txt`. Do not add `GROQ_API_KEY` to the app's shared secrets.
4. Share the deployed HTTPS URL. Each visitor enters and validates their own key; no deployment-side Groq key is
  required. Streamlit transports the entered key to the app server over HTTPS, and the server sends it to Groq
  over HTTPS.

The same app can run on another HTTPS-enabled Streamlit host with the dependencies in `requirements.txt` and
Python 3.10 or later.

The active key is held in Streamlit session state in server memory and is not written to disk or included in
downloads. Separate Streamlit sessions have separate state. Reset clears the key; a disconnected, expired, or
restarted session may lose it, in which case the user must enter it again. This is session isolation, not a vault:
the deployment operator and anyone with server-process access may be able to inspect process memory. Use a trusted
host, HTTPS, restricted operational access, and avoid request or state dumps that could capture secrets.
If a personal key was ever committed or exposed previously, revoke it in the Groq console and create a replacement.

## Groq limits (verified against Groq's docs on 9 Oct 2026)

These change over time. Check <https://console.groq.com/docs/speech-to-text> and
<https://console.groq.com/docs/rate-limits> for current values.

- Model `whisper-large-v3`. Supported file types: flac, mp3, mp4, mpeg, mpga, m4a, ogg, wav, webm.
- Maximum file size: **25 MB on the free tier** (100 MB on the Developer tier).
- Free-plan limits for Whisper models: **20 requests/minute, 2,000 requests/day, 7,200 audio seconds/hour,
  28,800 audio seconds/day**. Limits apply to your whole Groq organization.
- Requests under 10 seconds of audio count as 10 seconds toward usage.
- Audio is downsampled to 16 kHz mono; only the first audio track of a file is transcribed.
- This app sends each file in one request. For files above the limit, shrink them first with
  [ffmpeg](https://ffmpeg.org/): `ffmpeg -i input.mp3 -ar 16000 -ac 1 -map 0:a -c:a flac output.flac`
- Excel allows 32,767 characters per cell. A longer transcript continues on following rows with the same
  *Slide Name* value, so no text is lost.

## Troubleshooting

| Problem | Fix |
|---|---|
| "Enter your Groq API key first" | Enter and validate your own key in the Groq API Key section |
| "Groq rejected this API key" | Re-copy the key from the Groq console, confirm it is active, and validate again |
| "Groq rate limit reached" / "hourly/daily audio limit" | Wait and retry, process fewer files at a time, or use a Developer plan |
| "File is over the 25 MB limit" | Compress with the ffmpeg command above, or split the recording |
| "Unsupported file type" | Convert to mp3, wav, m4a, flac, ogg or webm |
| "Groq returned an empty transcription" | The file may be silent or contain no speech |
| "Could not connect to Groq" | Check your internet connection, VPN, or firewall |
| `streamlit` is not recognized | Activate the venv first: `.\venv\Scripts\Activate.ps1` |
| Port 8501 already in use | `streamlit run app.py --server.port 8502` |
| Excel cannot save to `output/` | Close any open copy of the file, or check folder permissions |

## Project structure

```
audio-transcription-project/
├── app.py                # Streamlit UI and workflow
├── transcription.py      # Groq calls, validation, retries
├── excel_generator.py    # Excel creation and formatting
├── config.py             # Settings and .env validation
├── requirements.txt
├── .env.example          # optional non-secret settings (copy to .env)
├── .gitignore
├── pytest.ini
├── README.md
├── output/               # generated Excel copies
└── tests/                # mocked unit tests + opt-in real API test
```

## Privacy

Audio is sent only to Groq's transcription API. User API keys are kept in Streamlit session state, not written to
Excel files, uploaded files, logs, or the repository. Transcripts and generated files may remain in the host's
`output/` folder, so deployment operators should account for that data separately. No other external service is
used by this app.
