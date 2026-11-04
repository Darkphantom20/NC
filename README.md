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

## AI drafting

Use **Draft with AI** to create editable text for the narrative report fields. Drafts are reviewed before they are applied to the form. The app supports Groq and Gemini through server-side API routes; it never sends provider keys to the browser.

Add one provider key to `.env.local` (do not commit this file):

```env
GROQ_API_KEY=your-rotated-groq-key
# or
GEMINI_API_KEY=your-rotated-gemini-key
```

If both keys are configured, Groq is used by default. Set `AI_PROVIDER=gemini` or `AI_PROVIDER=groq` to choose explicitly. Optional model overrides are `GROQ_MODEL` and `GEMINI_MODEL`. Restart the development server after changing environment variables. For Vercel, add the same variables under the project's Environment Variables settings and redeploy.

## Deploy to Vercel

1. Push this project to GitHub.
2. Import it into Vercel.
3. Use the default Next.js settings.

## Notes

Narratives are saved automatically in this browser using IndexedDB, which supports much more data than `localStorage` and is better suited to reports with attached images. Existing drafts from the earlier `localStorage` format are migrated automatically after the first successful save. Use **New narrative** to start a separate report, then use **Current narrative** to switch between reports. Browser-local drafts are not shared across browsers or devices, and the browser can still enforce storage limits, so generate/download a copy to keep a separate backup.

In **Page Layout Style**, use **Report Template** to choose either **Current System Layout** or **narrative-report.docx Layout**. The DOCX-based option uses the template's cover, acknowledgement, table of contents, organization sections, appendices, one-inch page margins, and Times New Roman structure while retaining the report data entered in the system.

In **Appendices → Daily Journal**, use **Import DOCX** to read journal tables with date and accomplishment columns. The import review matches any embedded documentation photo to the day row containing it; days without a photo stay unattached. Confirming replaces only that report's Daily Journal entries and keeps the selected journal output layout and other report sections unchanged. Hours are left blank when the source table has no hours column.
