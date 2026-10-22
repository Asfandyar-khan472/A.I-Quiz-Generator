# AI Test Paper Generator

Upload study material (PDF, DOCX or TXT notes/syllabus) or just type a topic, and get a complete,
formatted **test paper with an answer key** in seconds. Built with **Streamlit** and **Google Gemini (Flash)**.

## Features

- **Input:** upload PDF / DOCX / TXT, type a topic or chapter, or do both.
- **Paper settings:** subject, class/level, difficulty (Easy / Medium / Hard / Mixed).
- **Question types:** Multiple Choice, True/False, Short Answer, Long Answer. You choose how many of each and the marks per question; the total marks are calculated for you.
- **Optional header:** school name, paper title, time allowed, student instructions.
- **Answer key:** correct answers, explanations and marking schemes for written questions.
- **Regenerate one question** without redoing the whole paper.
- **Download** the paper and the answer key as **PDF** or **DOCX**.
- Handles empty, oversized, corrupt, password-protected and scanned files, invalid AI output, rate limits and API errors with clear messages.

## Project structure

```
ai-test-paper-generator/
├── app.py            # the whole Streamlit app
├── requirements.txt  # Python dependencies
└── README.md
```

## 1. Get a Gemini API key

1. Go to <https://aistudio.google.com/apikey> and sign in with a Google account.
2. Click **Create API key** and copy it. Keep it private.

## 2. Run locally

Requires Python 3.10 or newer.

```bash
# create and activate a virtual environment (recommended)
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate

# install dependencies
pip install -r requirements.txt

# provide your API key (choose ONE option)
# Option A: environment variable
export GEMINI_API_KEY="your-key-here"        # Windows PowerShell: $env:GEMINI_API_KEY="your-key-here"
# Option B: Streamlit secrets file  -> create .streamlit/secrets.toml containing:
#   GEMINI_API_KEY = "your-key-here"
# Option C: paste the key into the sidebar when the app is running

streamlit run app.py
```

The app opens at <http://localhost:8501>.

## 3. How to use

1. Upload your notes (or type a topic/chapter).
2. Fill in subject, class and difficulty.
3. Choose how many questions of each type and the marks for each.
4. Click **Generate test paper**.
5. Review the paper and answer key, regenerate any weak question, then download the PDF or DOCX.

## Configuration

| Setting | Where | Default |
|---|---|---|
| `GEMINI_API_KEY` (or `GOOGLE_API_KEY`) | Streamlit secrets, environment variable, or sidebar | none (required) |
| `GEMINI_MODEL` | Streamlit secrets or environment variable | `gemini-3.5-flash` |
| Model name | Sidebar text box | value of `GEMINI_MODEL` |

Google renames and retires models from time to time. If you see "model not found", change the model name in the sidebar to a current Flash model listed at <https://ai.google.dev/gemini-api/docs/models>.

Limits (editable at the top of `app.py`): 10 MB per file, first 80,000 characters of material used, 60 questions per paper.

## Known limitations

- **Scanned PDFs (images of text)** have no extractable text. Run OCR first, or paste the text into a TXT file.
- **PDF downloads use a Latin-1 font**, so Urdu, Arabic, Chinese and special math symbols show as `?` in the PDF. Use the **DOCX** download for those languages and symbols.
- AI can make mistakes. **Always review the questions and answer key before using a paper in a real exam.**
- Free-tier Gemini keys have rate limits. The app retries automatically, but heavy use may need a paid key.

## Push to GitHub

```bash
cd ai-test-paper-generator

# keep secrets and clutter out of the repo
cat > .gitignore <<'EOT'
venv/
__pycache__/
*.pyc
.env
.streamlit/secrets.toml
EOT

git init
git add .
git commit -m "Initial commit: AI Test Paper Generator"
git branch -M main
```

Create an **empty** repository on GitHub (<https://github.com/new>, no README or .gitignore), then:

```bash
git remote add origin https://github.com/<your-username>/ai-test-paper-generator.git
git push -u origin main
```

GitHub no longer accepts account passwords over HTTPS. When prompted, use a **Personal Access Token**
(GitHub > Settings > Developer settings > Personal access tokens), or sign in with GitHub CLI (`gh auth login`) or GitHub Desktop.

## Deploy on Streamlit Community Cloud (free)

1. Go to <https://share.streamlit.io> and sign in with GitHub.
2. Click **Create app** (or **New app**) and choose **Deploy a public app from GitHub**.
3. Select your repository, the `main` branch, and set **Main file path** to `app.py`.
4. Open **Advanced settings**, pick Python 3.11 or 3.12, and in **Secrets** paste:
   ```toml
   GEMINI_API_KEY = "your-key-here"
   ```
5. Click **Deploy**. The first build takes a few minutes. You get a public `*.streamlit.app` URL.

To change the key later: open your app > **Manage app** > **Settings** > **Secrets**.
Every `git push` to `main` redeploys the app automatically.

> If the app is public, anyone can use it and spend your Gemini quota. Restrict viewers in the app's sharing settings, or set a spending limit on your key.

## Troubleshooting

| Problem | Fix |
|---|---|
| `ModuleNotFoundError` on deploy | Make sure `requirements.txt` is in the repo root and spelled correctly. |
| "No API key yet" in the sidebar | Add `GEMINI_API_KEY` to Secrets (cloud) or set the env var / secrets file (local). |
| "Gemini rejected the API key" | Re-copy the key; confirm it was created in Google AI Studio. |
| "Model was not found" | Change the model name in the sidebar. |
| "Rate limit or quota reached" | Wait a minute, request fewer questions, or use a key with higher quota. |
| Fewer questions than requested | The AI returned some invalid items. Click Generate again or regenerate single questions. |

## License

MIT. Use and modify freely.
