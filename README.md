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

## Groq AI text generation

Set `GROQ_API_KEY` as a server-only environment variable.

```powershell
$env:GROQ_API_KEY="your_groq_api_key"
npm run dev
```

For Vercel, add `GROQ_API_KEY` under **Project Settings → Environment Variables**, then redeploy. Optional `GROQ_MODEL` defaults to `openai/gpt-oss-20b`. Never use a `NEXT_PUBLIC_` prefix for the key.

## Gemini backup text generation

Set `GEMINI_API_KEY` as another server-only environment variable. When Groq is unavailable, rate-limited, or returns no text, the API automatically retries with Gemini.

```powershell
$env:GEMINI_API_KEY="your_gemini_api_key"
npm run dev
```

For Vercel, add both `GROQ_API_KEY` and `GEMINI_API_KEY` under **Project Settings → Environment Variables**, then redeploy. Optional `GEMINI_MODEL` defaults to `gemini-flash-latest`. Treat both keys as secrets and rotate any key that has been shared publicly.

## Notes

Narratives are saved automatically in this browser using IndexedDB, which supports much more data than `localStorage` and is better suited to reports with attached images. Existing drafts from the earlier `localStorage` format are migrated automatically after the first successful save. Use **New narrative** to start a separate report, then use **Current narrative** to switch between reports. Browser-local drafts are not shared across browsers or devices, and the browser can still enforce storage limits, so generate/download a copy to keep a separate backup.

In **Page Layout Style**, use **Report Template** to choose either **Current System Layout** or **narrative-report.docx Layout**. The DOCX-based option uses the template's cover, acknowledgement, table of contents, organization sections, appendices, one-inch page margins, and Times New Roman structure while retaining the report data entered in the system.

In **Appendices → Daily Journal**, use **Import DOCX** to read journal tables with date and accomplishment columns. The import review matches any embedded documentation photo to the day row containing it; days without a photo stay unattached. Confirming replaces only that report's Daily Journal entries and keeps the selected journal output layout and other report sections unchanged. Hours are left blank when the source table has no hours column.
