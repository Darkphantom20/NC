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

Narratives are saved automatically in the current browser. Use **New narrative** to start a separate report, then use **Current narrative** to switch between saved reports. Existing single-report data is migrated automatically the first time the app loads the new version. Browser-local drafts are not shared across browsers or devices, so generate/download a copy to keep a separate backup.
