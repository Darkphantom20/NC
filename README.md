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

In **Page Layout Style**, use **Report Template** to choose either **Current System Layout** or **narrative-report.docx Layout**. The DOCX-based option uses the template's cover, acknowledgement, table of contents, organization sections, appendices, one-inch page margins, and Times New Roman structure while retaining the report data entered in the system.

In **Appendices → Daily Journal**, use **Import DOCX** to read journal tables with date and accomplishment columns. The import review matches any embedded documentation photo to the day row containing it; days without a photo stay unattached. Confirming replaces only that report's Daily Journal entries and keeps the selected journal output layout and other report sections unchanged. Hours are left blank when the source table has no hours column.

Use **Import a full narrative DOCX** to read report text from Word paragraphs and tables, including labeled cover-page details and recognized section headings. Review and apply the fields found; only matching fields are replaced. The generated document uses the selected system template and does not preserve the source DOCX's exact styling, page breaks, or images.

In the Daily Journal, use **Generate with AI** below a day's accomplishment to draft a concise entry from that day's attached photo and existing notes. The draft opens in the editor for review before it is saved. Photo generation requires a server-side Gemini API key or Cloudflare Workers AI credentials; Gemini is tried first unless `AI_PROVIDER=cloudflare`, and the Cloudflare fallback uses the vision-capable `@cf/meta/llama-4-scout-17b-16e-instruct` model. Add the provider credentials to local `.env.local` and to the production hosting environment; never expose them in client-side variables.

```env
GEMINI_API_KEY=your-gemini-api-key
# or / as a fallback
CLOUDFLARE_ACCOUNT_ID=your-cloudflare-account-id
CLOUDFLARE_API_TOKEN=your-cloudflare-api-token
AI_PROVIDER=cloudflare
```
