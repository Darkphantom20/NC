# JRMSU Narrative Report Generator

This project is a Vercel-ready web app version of the narrative report generator.

## Stack
- Next.js
- TypeScript
- Tailwind CSS

## Run locally

```bash
npm install
npm run dev
```

## Deploy to Vercel

1. Push this project to GitHub.
2. Import it into Vercel.
3. Use the default Next.js settings.

## Notes

Narratives are saved automatically in this browser using IndexedDB, which supports much more data than `localStorage` and is better suited to reports with attached images. Existing drafts from the earlier `localStorage` format are migrated automatically after the first successful save. Use **New narrative** to start a separate report, then use **Current narrative** to switch between reports. Browser-local drafts are not shared across browsers or devices, and the browser can still enforce storage limits, so generate/download a copy to keep a separate backup.
