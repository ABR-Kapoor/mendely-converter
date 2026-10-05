# Mendeley Converter

A local Streamlit app for analyzing IEEE-style citations and generating a Mendeley Cite v3-compatible DOCX structure.

## Local development

```powershell
cd "C:\Users\NITR\Music\Abeer-researches\code-show\mendely-converter"
.\venv\Scripts\streamlit.exe run app.py
```

## Vercel-compatible deployment

This repository includes a minimal Vercel-compatible Python API wrapper under `api/index.py` and a `vercel.json` config.

Important:
- Vercel does not natively host Streamlit apps directly.
- The Streamlit UI remains for local usage.
- The repo is now GitHub-ready and includes a Vercel-compatible API entrypoint.

## GitHub

Commit this folder to a GitHub repository and deploy the project with Vercel using the included `vercel.json`.
