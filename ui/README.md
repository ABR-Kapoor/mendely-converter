UI subfolder

This is a static frontend intended to be deployed separately (e.g., Vercel static site) and to call the Flask API at `/api/convert`.

How to run locally:

1. Simple: open `ui/index.html` in a browser (CORS may block the API calls — best test by deploying or running a local server).

2. Quick local server (from repo root):

```bash
python -m http.server --directory ui 8080
# then open http://localhost:8080
```

Design notes:
- Light theme inspired by makro.framer.website (fonts, spacing, accents).
- Uses Inter from Google Fonts.
- Upload form posts to `/api/convert` (the Flask API in this repo).

To deploy to Vercel:
- Add the `ui` folder as a separate project in Vercel, or include rewrite rules to serve static `ui` content.
