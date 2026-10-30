import fs from 'fs';
import path from 'path';
import { createClient } from '@supabase/supabase-js';
import { Document, Packer, Paragraph, TextRun, AlignmentType, ImageRun, Header, Footer, Table, TableRow, TableCell, WidthType, BorderStyle, TableLayoutType } from 'docx';

export const runtime = 'nodejs';

const OUTPUT_FILE_NAME = process.env.OUTPUT_FILE_NAME || 'JRMSU_Narrative_Report.docx';
const HEADER_IMAGE_PATH = path.join(process.cwd(), 'public', 'Picture1.png');
const HEADER_IMAGE = fs.existsSync(HEADER_IMAGE_PATH) ? fs.readFileSync(HEADER_IMAGE_PATH) : null;
const SCANNED_TEMPLATE_HEADER_PATH = path.join(process.cwd(), 'public', 'Untitled.png');
const SCANNED_TEMPLATE_HEADER = fs.existsSync(SCANNED_TEMPLATE_HEADER_PATH) ? fs.readFileSync(SCANNED_TEMPLATE_HEADER_PATH) : null;

interface AppendixImage {
  base64: string;
  detail?: string;
}

interface DailyActivity {
  day: string;
  date: string;
  accomplishment: string;
  hours: number | string;
}

interface DailyJournalWeek {
  weekNumber: number;
  activities: DailyActivity[];
  totalHours: number | string;
  narrative?: string;
  images?: (string | AppendixImage)[]; // For backwards compatibility
  reportLayoutImages?: (string | AppendixImage)[]; // Report layout only
  currentLayoutImages?: (string | AppendixImage)[]; // Current layout only
}

interface ReportFooterData {
  preparedByLabel?: string;
  preparedByName?: string;
  checkedByLabel?: string;
  checkedByName?: string;
  checkedByRole?: string;
  officeInCharge?: string;
  dateLabel?: string;
  dateValue?: string;
}

interface AppendicesData {
  dailyJournalLayout?: 'current' | 'report' | 'no-photo';
  reportFooter?: ReportFooterData;
  dailyJournal?: DailyJournalWeek[];
  certParticipation?: string;
  primeNarrative?: string;
  primeReportImages?: (string | AppendixImage)[];
  resume?: string;
  grades?: string;
  medicalWaiver?: string;
  letterAcceptance?: string;
  dtr?: string;
  ratingSheet?: string;
  certCompletion?: string;
}

interface SectionData {
  title: string;
  content: any;
  isBullet?: boolean;
}

function buildHeaderImage(headerImage: Buffer | null = HEADER_IMAGE, width = 624, height = 125) {
  if (!headerImage) return [];
  return [
    new ImageRun({
      data: headerImage,
      type: 'png',
      transformation: { width, height },
    }),
  ];
}

const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;
const hasSupabaseConfig = Boolean(supabaseUrl && supabaseKey);

function ensureArray(val: any): string[] {
  if (Array.isArray(val)) return val.map(v => String(v).trim()).filter(Boolean);
  if (typeof val === 'string' && val.trim().length > 0) {
    return val.split(/\r?\n/).map(v => v.trim()).filter(Boolean);
  }
  return [];
}

function parseBase64Image(base64String: string): Buffer | null {
  if (!base64String || typeof base64String !== 'string') return null;
  if (!base64String.startsWith('data:image/')) return null;

  try {
    const cleanBase64 = base64String.replace(/^data:image\/\w+;base64,/, "");
    return Buffer.from(cleanBase64, 'base64');
  } catch (err) {
    console.error("Failed to parse image", err);
    return null;
  }
}

export async function POST(request: Request) {
  try {
    const data = await request.json();

    if (!data || typeof data !== 'object') {
      return new Response(JSON.stringify({ error: 'Invalid payload' }), { status: 400 });
    }

    const isPreviewRequest = request.headers.get('x-report-preview') === 'true';

    if (hasSupabaseConfig && !isPreviewRequest) {
      const supabase = createClient(supabaseUrl!, supabaseKey!);
      const payload = {
        student_name: data.studentName || null,
        appendices: JSON.stringify(data.appendices) || null,
      };
      await supabase.from('reports').insert([payload]);
    }

    const reportFooterChildren = data.appendices?.reportFooter
      ? buildReportFooterBlock(data.appendices?.reportFooter)
      : [];
    const compactSectionLayout = data.sectionLayout === 'compact';
    const acknowledgementLineSpacing = 360;

    const documentSections: any[] = [
      {
        properties: {
          page: {
            size: { width: 12240, height: 20160 },
            margin: { top: 1440, bottom: 1440, left: 1440, right: 1440, header: 0, footer: 0 },
          },
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({ alignment: AlignmentType.CENTER, children: buildHeaderImage() }),
            ],
          }),
        },
        children: [
          ...buildCoverPage(data),
          new Paragraph({ pageBreakBefore: true }),
          ...buildAcknowledgementPage(data, compactSectionLayout, acknowledgementLineSpacing),
          new Paragraph({ pageBreakBefore: true }),
          ...buildTableOfContentsPage(data, compactSectionLayout, acknowledgementLineSpacing),
          new Paragraph({ pageBreakBefore: true }),
          ...buildSectionPage('INTRODUCTION', [
            { title: 'Background of the Organization', content: data.background },
            { title: 'Vision', content: data.vision },
            { title: 'Mission', content: data.mission },
            { title: 'Objectives', content: data.objectives, isBullet: true },
            { title: 'Core Values', content: data.coreValues, isBullet: true },
            { title: 'Products and Services Offered', content: data.services },
          ], compactSectionLayout, acknowledgementLineSpacing, data.organizationStructureImage),
          new Paragraph({ pageBreakBefore: !compactSectionLayout }),
          ...buildSectionPage('ORGANIZATION / COMPANY ANALYSIS', [
            { title: 'Strengths', content: data.strengths, isBullet: true },
            { title: 'Weaknesses', content: data.weaknesses, isBullet: true },
            { title: 'Opportunities', content: data.opportunities, isBullet: true },
            { title: 'Threats', content: data.threats, isBullet: true },
            { title: 'Recommendations for Improvement', content: data.recommendations, isBullet: true },
          ], compactSectionLayout, acknowledgementLineSpacing),
          new Paragraph({ pageBreakBefore: !compactSectionLayout }),
          ...buildSectionPage('TASKS AND DUTIES', [
            { title: 'Assigned Tasks and Responsibilities', content: data.tasks, isBullet: true },
            { title: 'Duties and Procedures Conformed', content: data.procedures, isBullet: true },
          ], compactSectionLayout, acknowledgementLineSpacing),
          new Paragraph({ pageBreakBefore: !compactSectionLayout }),
          ...buildSectionPage('CASE ANALYSIS', [
            { title: 'Issue / Problem 1', content: data.issue1 },
            { title: 'Strategy/Action Undertaken for Problem 1', content: data.issue1Action },
            { title: 'Issue / Problem 2', content: data.issue2 },
            { title: 'Strategy/Action Undertaken for Problem 2', content: data.issue2Action },
            { title: 'Lessons Learned from the Situations', content: data.lessons },
          ], compactSectionLayout, acknowledgementLineSpacing),
          new Paragraph({ pageBreakBefore: !compactSectionLayout }),
          ...buildSectionPage('REFLECTIONS', [
            { title: 'Self-Evaluation', content: data.selfEvaluation },
            { title: 'Relevancy of the Organization', content: data.relevancy },
          ], compactSectionLayout, acknowledgementLineSpacing),
          ...(data.schoolPlacement?.enabled
            ? [
              new Paragraph({ pageBreakBefore: true }),
              ...buildSchoolPlacementPages(data.schoolPlacement, compactSectionLayout, acknowledgementLineSpacing),
            ]
            : []),
        ],
      },
    ];

    const dailyJournalContent = buildDailyJournalAppendixPage(data.appendices);
    if (dailyJournalContent.length > 0) {
      documentSections.push({
        properties: {
          page: {
            size: { width: 12240, height: 20160 },
            margin: { top: 1440, bottom: 1440, left: 1440, right: 1440, header: 0, footer: 0 },
          },
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({ alignment: AlignmentType.CENTER, children: buildHeaderImage() }),
            ],
          }),
        },
        footers: reportFooterChildren.length > 0 ? {
          default: new Footer({
            children: reportFooterChildren,
          }),
        } : undefined,
        children: [
          ...(!compactSectionLayout ? [new Paragraph({ pageBreakBefore: true })] : []),
          ...dailyJournalContent,
        ],
      });
    }

    const remainingAppendixContent = buildOtherAppendixPages(data.appendices);
    if (remainingAppendixContent.length > 0) {
      documentSections.push({
        properties: {
          page: {
            size: { width: 12240, height: 20160 },
            margin: { top: 1440, bottom: 1440, left: 1440, right: 1440, header: 0, footer: 0 },
          },
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({ alignment: AlignmentType.CENTER, children: buildHeaderImage() }),
            ],
          }),
        },
        children: [
          ...(!compactSectionLayout ? [new Paragraph({ pageBreakBefore: true })] : []),
          ...remainingAppendixContent,
        ],
      });
    }

    if (data.reportTemplate === 'docx-template') {
      documentSections.splice(
        0,
        documentSections.length,
        ...buildDocxTemplateSections(data),
      );
    }

    if (data.reportTemplate === 'scanned-template') {
      documentSections.splice(
        0,
        documentSections.length,
        ...buildDocxTemplateSections(data, SCANNED_TEMPLATE_HEADER, true),
      );
    }

    const doc = new Document({
      sections: documentSections,
    });

    const buffer = await Packer.toBuffer(doc);
    const docBuffer = Buffer.from(buffer);

    return new Response(docBuffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="${OUTPUT_FILE_NAME}"`,
      },
    });
  } catch (error) {
    console.error(error);
    return new Response(JSON.stringify({ error: 'Something went wrong while generating the report.' }), { status: 500 });
  }
}

function templateParagraph(text: string, options: {
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
  bold?: boolean;
  size?: number;
  before?: number;
  after?: number;
  pageBreakBefore?: boolean;
  justified?: boolean;
  bullet?: boolean;
} = {}): Paragraph[] {
  const paragraph = new Paragraph({
    alignment: options.alignment ?? (options.justified ? AlignmentType.JUSTIFIED : AlignmentType.LEFT),
    pageBreakBefore: options.pageBreakBefore,
    spacing: {
      before: options.before ?? 0,
      after: options.after ?? 0,
      line: 360,
      lineRule: 'auto',
    },
    indent: options.bullet ? { left: 720 } : undefined,
    bullet: options.bullet ? { level: 0 } : undefined,
    children: [new TextRun({
      text,
      bold: options.bold ?? false,
      size: options.size ?? 24,
      font: 'Times New Roman',
    })],
  });
  return [paragraph];
}

function buildTemplateCoverPage(data: any): Paragraph[] {
  const placementOrganizations = [
    data.trainingOrganization,
    data.schoolPlacement?.enabled ? data.schoolPlacement.trainingOrganization : '',
  ].filter(Boolean);
  const placementLocations = [
    data.trainingLocation,
    data.schoolPlacement?.enabled ? data.schoolPlacement.trainingLocation : '',
  ].filter(Boolean);

  const centered = (text: string, options: {
    before?: number;
    after?: number;
    bold?: boolean;
    size?: number;
    line?: number;
  } = {}) => new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: {
      before: options.before ?? 0,
      after: options.after ?? 360,
      line: options.line ?? 360,
      lineRule: 'auto',
    },
    children: [new TextRun({
      text,
      bold: options.bold ?? false,
      size: options.size ?? 32,
      font: 'Times New Roman',
    })],
  });

  return [
    centered('A Narrative Report on the', { before: 1440, after: 120, size: 32 }),
    centered('On-the-Job Training conducted at', { after: 240, size: 32 }),
    centered(placementOrganizations.join(' and '), { bold: true, after: 120, size: 32 }),
    centered(placementLocations.join(' and '), { after: 1440, size: 32 }),
    centered('Presented to the faculty of', { after: 120, size: 32 }),
    centered(data.collegeFaculty || '', { after: 1440, size: 32 }),
    centered('In partial fulfillment of the requirements for the degree of', { after: 120, size: 32 }),
    centered(data.degreeProgram || '', { after: 2160, size: 32 }),
    centered('Training Duration:', { after: 120, size: 32 }),
    centered(data.trainingDuration || '', { after: 1200, size: 32 }),
    centered('Submitted by:', { after: 120, size: 32 }),
    centered(data.studentName || '', { bold: true, after: 120, size: 32 }),
    centered(`${data.degreeProgram || ''}${data.studentYear ? ` - ${data.studentYear}` : ''}`, { after: 1440, size: 32 }),
    centered('Submitted to:', { after: 120, size: 32 }),
    centered(data.submittedToName || '', { bold: true, after: 120, size: 32 }),
    centered(data.submittedToTitle || '', { after: 0, size: 32 }),
  ];
}

function buildTemplateTableOfContents(data: any): any[] {
  const schoolEnabled = data.schoolPlacement?.enabled;
  const rows: [string, string][] = [
    ['Acknowledgement', 'i'],
    ['Table of Contents', 'ii'],
    ['List of Table/Figures', 'iii'],
    ['I. ' + String(data.trainingOrganization || 'TRAINING ORGANIZATION').toUpperCase(), ''],
    ['Introduction', '1'],
    ['  Background of the Organization', ''],
    ['  Vision', ''],
    ['  Mission', ''],
    ['  Objectives', ''],
    ['  Core Values', ''],
    ['  Products and Services Offered', ''],
    ['Organization / Company Analysis', '4'],
    ['Tasks and Duties', '6'],
    ['Case Analysis', '7'],
    ['Reflections', '8'],
  ];
  if (schoolEnabled) {
    rows.push(
      ['II. ' + String(data.schoolPlacement?.trainingOrganization || 'SECOND OJT PLACEMENT').toUpperCase(), ''],
      ['Introduction', '9'],
      ['  Background of the Organization', ''],
      ['  Vision', ''],
      ['  Mission', ''],
      ['  Objectives', ''],
      ['  Core Values', ''],
      ['  Products and Services Offered', ''],
      ['Organization / Company Analysis', '12'],
      ['Tasks and Duties', '14'],
      ['Case Analysis', '15'],
      ['Reflections', '16'],
      ['Appendices', '17'],
    );
  } else {
    rows.push(['Appendices', '9']);
  }

  return rows.map(([title, page]) => new Paragraph({
    spacing: { after: 60, line: 360, lineRule: 'auto' },
    children: [
      new TextRun({ text: `${title}${page ? ' ' + '-'.repeat(Math.max(8, 82 - title.length)) : ''}`, size: 24, font: 'Times New Roman', bold: /^(I\.|II\.|Organization\/Company Analysis|Tasks and Duties|Case Analysis|Reflections)/.test(title) }),
      ...(page ? [
        new TextRun({ text: `  ${page}`, size: 24, font: 'Times New Roman' }),
      ] : []),
    ],
  }));
}

function buildTemplateDailyJournal(appendices: AppendicesData | undefined): (Paragraph | Table)[] {
  if (!appendices?.dailyJournal?.length) return [];
  const children: (Paragraph | Table)[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 240, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: 'APPENDICES', bold: true, size: 24, font: 'Times New Roman' })],
    }),
  ];
  const layout = appendices.dailyJournalLayout === 'no-photo'
    ? 'no-photo'
    : appendices.dailyJournalLayout === 'report' ? 'report' : 'current';

  appendices.dailyJournal.forEach((week, index) => {
    if (layout === 'report') {
      const reportImages = week.reportLayoutImages || week.images || [];
      children.push(
        new Paragraph({
          pageBreakBefore: index > 0,
          spacing: { before: 120, after: 120, line: 360, lineRule: 'auto' },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: `WEEK ${week.weekNumber}`, bold: true, size: 24, font: 'Times New Roman' })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 60, after: 120, line: 360, lineRule: 'auto' },
          children: [new TextRun({ text: 'WEEKLY ACCOMPLISHMENT REPORT', bold: true, size: 24, font: 'Times New Roman' })],
        }),
        buildDailyReportTable(week.activities, reportImages),
      );
      return;
    }

    children.push(
      new Paragraph({
        pageBreakBefore: index > 0,
        spacing: { before: 120, after: 120, line: 360, lineRule: 'auto' },
        children: [new TextRun({ text: `Week ${week.weekNumber}`, bold: true, size: 24, font: 'Times New Roman' })],
      }),
      buildCurrentWeeklyTable(week.activities, week.totalHours),
    );

    if (layout === 'current') {
      children.push(
        new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 120, line: 360, lineRule: 'auto' },
        children: [new TextRun({ text: `Table ${week.weekNumber}: Journal Week ${week.weekNumber}`, bold: true, size: 24, font: 'Times New Roman' })],
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 120, line: 360, lineRule: 'auto' },
        children: [new TextRun({ text: week.narrative || 'No narrative provided.', size: 24, font: 'Times New Roman' })],
      }),
      );
    }

    if (layout === 'current') {
      const images = (week.currentLayoutImages || week.images || []).filter((image): image is string | AppendixImage => Boolean(image));
      if (images.length > 0) {
        children.push(buildImageGridTable(images));
        images.forEach((image, imageIndex) => {
          const detail = typeof image === 'string' ? '' : image.detail?.trim();
          children.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 60, after: 120, line: 360, lineRule: 'auto' },
            children: [new TextRun({
              text: `Figure ${imageIndex + 1}${detail ? `: ${detail}` : ''}`,
              size: 24,
              font: 'Times New Roman',
            })],
          }));
        });
      }
    }
  });

  return children;
}

function buildTemplateOrganizationPages(label: string, source: any, romanLabel = '', figureNumber = 1): Paragraph[] {
  const children: Paragraph[] = [
    new Paragraph({
      pageBreakBefore: true,
      alignment: AlignmentType.CENTER,
      spacing: { after: 120, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: `${romanLabel}${romanLabel ? '. ' : ''}${String(label || 'TRAINING ORGANIZATION').toUpperCase()}`, bold: true, size: 24, font: 'Times New Roman' })],
    }),
  ];

  const addPageHeading = (text: string) => {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 180, line: 360, lineRule: 'auto' },
        children: [new TextRun({ text, bold: true, size: 24, font: 'Times New Roman' })],
      }),
    );
  };
  const addSubheading = (text: string) => {
    children.push(
      new Paragraph({
        spacing: { before: 60, after: 60, line: 360, lineRule: 'auto' },
        children: [new TextRun({ text: `${text}:`, bold: true, size: 24, font: 'Times New Roman' })],
      }),
    );
  };
  const addText = (text: string, bullet = false) => {
    const lines = ensureArray(text);
    if (!lines.length) return;
    lines.forEach((line) => children.push(
      new Paragraph({
        alignment: bullet ? AlignmentType.LEFT : AlignmentType.JUSTIFIED,
        indent: bullet ? { left: 720 } : undefined,
        bullet: bullet ? { level: 0 } : undefined,
        spacing: { before: 0, after: 0, line: 360, lineRule: 'auto' },
        children: [new TextRun({ text: line, size: 24, font: 'Times New Roman' })],
      }),
    ));
  };

  addPageHeading('INTRODUCTION');
  addSubheading('Background of the Organization');
  addText(source.background);
  addSubheading('Vision');
  addText(source.vision);
  addSubheading('Mission');
  addText(source.mission);
  addSubheading('Objectives');
  addText(source.objectives, true);
  addSubheading('Core Values');
  addText(source.coreValues, true);
  addSubheading('Products and Services Offered');
  addText(source.services, true);

  const imageBuffer = parseBase64Image(source.organizationStructureImage || '');
  if (imageBuffer) {
    children.push(
      ...templateParagraph('Organization Structure', { alignment: AlignmentType.CENTER, bold: true, before: 120 }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new ImageRun({ type: 'png', data: imageBuffer, transformation: { width: 520, height: 360 } })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 120, line: 360, lineRule: 'auto' },
        children: [new TextRun({ text: `Figure ${figureNumber}: Organizational Chart`, size: 24, font: 'Times New Roman' })],
      }),
    );
  }

  children.push(
    new Paragraph({ pageBreakBefore: true }),
    ...templateParagraph('ORGANIZATION / COMPANY ANALYSIS', { alignment: AlignmentType.CENTER, bold: true }),
  );
  addSubheading('Strengths');
  addText(source.strengths, true);
  addSubheading('Weaknesses');
  addText(source.weaknesses, true);
  addSubheading('Opportunities');
  addText(source.opportunities, true);
  addSubheading('Threats');
  addText(source.threats, true);
  addSubheading('Recommendations for Improvement');
  addText(source.recommendations, true);

  children.push(
    new Paragraph({ pageBreakBefore: true }),
    ...templateParagraph('TASKS AND DUTIES', { alignment: AlignmentType.CENTER, bold: true }),
  );
  addSubheading('Assigned Tasks and Responsibilities');
  addText(source.tasks, true);
  addSubheading('Duties and Procedures Conformed');
  addText(source.procedures, true);

  children.push(
    new Paragraph({ pageBreakBefore: true }),
    ...templateParagraph('CASE ANALYSIS', { alignment: AlignmentType.CENTER, bold: true }),
  );
  addSubheading('Issue / Problem 1');
  addText(source.issue1);
  addSubheading('Strategy/Action Undertaken for Problem 1');
  addText(source.issue1Action);
  addSubheading('Issue / Problem 2');
  addText(source.issue2);
  addSubheading('Strategy/Action Undertaken for Problem 2');
  addText(source.issue2Action);
  addSubheading('Lessons Learned from the Situations');
  addText(source.lessons);

  children.push(
    new Paragraph({ pageBreakBefore: true }),
    ...templateParagraph('REFLECTIONS', { alignment: AlignmentType.CENTER, bold: true }),
  );
  addSubheading('Self-Evaluation');
  addText(source.selfEvaluation);
  addSubheading('Relevancy of the Organization');
  addText(source.relevancy);

  return children;
}

function buildDocxTemplateSections(
  data: any,
  headerImage: Buffer | null = HEADER_IMAGE,
  scannedTemplate = false,
): any[] {
  const children: (Paragraph | Table)[] = [
    ...buildTemplateCoverPage(data),
    new Paragraph({ pageBreakBefore: true }),
    ...buildAcknowledgementPage(data, false, 360),
    new Paragraph({ pageBreakBefore: true }),
    ...templateParagraph('TABLE OF CONTENTS', { alignment: AlignmentType.CENTER, bold: true, after: 180 }),
    ...buildTemplateTableOfContents(data),
    new Paragraph({ pageBreakBefore: true }),
    ...templateParagraph('LIST OF TABLES / FIGURES', { alignment: AlignmentType.CENTER, bold: true, after: 180 }),
    ...templateParagraph('LIST OF TABLES', { bold: true, after: 60 }),
  ];

  (data.appendices?.dailyJournal || []).forEach((week: DailyJournalWeek, index: number) => {
    children.push(...templateParagraph(`Table ${index + 1}: Journal Week ${week.weekNumber}`, {}));
  });

  children.push(
    ...templateParagraph('LIST OF FIGURES', { bold: true, before: 120, after: 60 }),
  );

  if (data.organizationStructureImage) {
    children.push(...templateParagraph('Figure 1: Organizational Chart', {}));
  }
  if (data.schoolPlacement?.enabled && data.schoolPlacement.organizationStructureImage) {
    children.push(...templateParagraph('Figure 2: School Placement Organizational Chart', {}));
  }
  const hasJournalPhotos = (data.appendices?.dailyJournal || []).some((week: DailyJournalWeek) =>
    (week.currentLayoutImages || week.images || []).some((image) => Boolean(image)),
  );
  if (hasJournalPhotos) {
    children.push(...templateParagraph('Figure 3: Weekly documentation photos', {}));
  }
  if (data.appendices?.primeReportImages?.length) {
    children.push(...templateParagraph('Figure 4: PRIME report photos', {}));
  }

  children.push(...buildTemplateOrganizationPages(data.trainingOrganization, data, 'I', 1));
  if (data.schoolPlacement?.enabled) {
    children.push(...buildTemplateOrganizationPages(data.schoolPlacement.trainingOrganization || 'SCHOOL OJT / SECOND PLACEMENT', data.schoolPlacement, 'II', 2));
  }

  const journalChildren = buildTemplateDailyJournal(data.appendices);
  if (journalChildren.length > 0) {
    children.push(new Paragraph({ pageBreakBefore: true }), ...journalChildren);
  }
  const otherAppendices = buildOtherAppendixPages(data.appendices);
  if (otherAppendices.length > 0) {
    children.push(...otherAppendices);
  }

  return [{
    properties: {
      page: {
        size: scannedTemplate
          ? { width: 12240, height: 20160 }
          : { width: 12242, height: 20163 },
        margin: scannedTemplate
          ? { top: 762000, bottom: 914400, left: 914400, right: 914400, header: 449580, footer: 449580 }
          : { top: 1440, bottom: 1440, left: 1440, right: 1440, header: 709, footer: 709 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: buildHeaderImage(
            headerImage,
            scannedTemplate ? 624 : 624,
            scannedTemplate ? 234 : 125,
          ),
        })],
      }),
    },
    children,
  }];
}

function buildCoverPage(data: any): Paragraph[] {
  const org = [data.trainingOrganization, data.schoolPlacement?.enabled ? data.schoolPlacement.trainingOrganization : '']
    .filter(Boolean)
    .join(' and ');
  const location = [data.trainingLocation, data.schoolPlacement?.enabled ? data.schoolPlacement.trainingLocation : '']
    .filter(Boolean)
    .join(' and ');
  const faculty = data.collegeFaculty || '';
  const degree = data.degreeProgram || '';
  const student = data.studentName || '';
  const studentDegree = data.degreeProgram || '';
  const adviserName = data.submittedToName || '';
  const adviserTitle = data.submittedToTitle || '';
  const FRONT_PAGE_TEXT_SIZE = 20;

  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 1440, after: 360 },
      children: [new TextRun({ text: 'A Narrative Report on the', size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 360, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: 'On-the-Job Training conducted at', size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 180, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: org, size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 1440, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: location, size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 360, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: 'Presented to the faculty of', size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 1440, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: faculty, size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 360, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: 'In partial fulfillment of the requirements for the degree of', size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 2160, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: degree, size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 180, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: 'Submitted by:', size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 180, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: student, size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 1080, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: studentDegree, size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 180, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: 'Submitted to:', size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 180, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: adviserName, size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 0, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: adviserTitle, size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
  ];
}

function buildAcknowledgementPage(data: any, compact = false, lineSpacing = 240): Paragraph[] {
  const student = data.studentName || '';
  const degree = data.degreeProgram || '';
  const university = 'Jose Rizal Memorial State University';
  const paragraphs: Paragraph[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
      children: [new TextRun({ text: 'ACKNOWLEDGEMENT', bold: true, size: 24, font: 'Times New Roman' })],
    }),
  ];

  const rawContent = data.acknowledgement || '';
  const textParagraphs = ensureArray(rawContent);

  if (textParagraphs.length > 0) {
    textParagraphs.forEach((pText) => {
      paragraphs.push(
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          indent: { firstLine: 720 },
          spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
          children: [new TextRun({ text: pText, size: 24, font: 'Times New Roman' })],
        })
      );
    });
  } else {
    paragraphs.push(
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        indent: { firstLine: 720 },
        spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
        children: [new TextRun({ text: 'No acknowledgement provided.', size: 24, font: 'Times New Roman' })],
      })
    );
  }

  // Signature Block matching the uploaded image format exactly
  paragraphs.push(
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
      children: [new TextRun({ text: student, bold: true, size: 24, font: 'Times New Roman' })],
    }),
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
      children: [new TextRun({ text: degree, size: 24, font: 'Times New Roman' })],
    }),
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
      children: [new TextRun({ text: university, size: 24, font: 'Times New Roman' })],
    })
  );

  return paragraphs;
}

function normalizeTOCKey(value: string): string {
  return value
    .replace(/^\s*\d+(?:\.\d+)*[.)]?\s*/, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\b(and|of|the|for|from|to|a)\b/g, ' ')
    .trim();
}

const ESTIMATED_PAGE_LINES = 44;
const BODY_CHARACTERS_PER_LINE = 84;

function estimateWrappedLines(text: string, charactersPerLine = BODY_CHARACTERS_PER_LINE): number {
  return text.split(/\r?\n/).reduce((total, line) => {
    return total + Math.max(1, Math.ceil(line.length / charactersPerLine));
  }, 0);
}

function estimateSectionLineCount(sectionData: SectionData[], hasImage = false): number {
  let totalLines = 1;

  sectionData.forEach((section) => {
    const text = typeof section.content === 'string'
      ? section.content
      : Array.isArray(section.content)
        ? section.content.join('\n')
        : String(section.content ?? '');
    const lines = ensureArray(text);

    totalLines += estimateWrappedLines(`${section.title}:`);
    if (section.isBullet) {
      totalLines += lines.reduce((count, line) => count + estimateWrappedLines(line), 0);
    } else if (lines.length > 0) {
      totalLines += estimateWrappedLines(lines.join(' '));
    }
  });

  if (hasImage) totalLines += 17;
  return totalLines;
}

function estimateSectionPageCount(sectionData: SectionData[], hasImage = false): number {
  return Math.max(1, Math.ceil(estimateSectionLineCount(sectionData, hasImage) / ESTIMATED_PAGE_LINES));
}

function addSubsectionPageMap(
  pageMap: Record<string, number>,
  sectionNumber: string,
  sectionData: SectionData[],
  startPage: number,
  compact: boolean,
) {
  let lineCursor = 1;

  sectionData.forEach((section, index) => {
    pageMap[`${sectionNumber}.${index + 1}`] = compact
      ? startPage + Math.floor(lineCursor / ESTIMATED_PAGE_LINES)
      : startPage;
    lineCursor += estimateSectionLineCount([section]);
  });
}

function estimateAcknowledgementLineCount(data: any): number {
  const paragraphs = ensureArray(data.acknowledgement || '');
  return 4 + paragraphs.reduce((total, paragraph) => total + estimateWrappedLines(paragraph), 0);
}

function estimateTOCLineCount(data: any): number {
  const entries = ensureArray(data.tableOfContents || '');
  const baseEntryCount = entries.filter((entry) => !isSchoolTOCLine(entry)).length || 26;
  return 1 + baseEntryCount + (data.schoolPlacement?.enabled ? 25 : 0);
}

function isSchoolTOCLine(title: string): boolean {
  return /school ojt/i.test(title)
    || /^(?:6|7|8|9|10)\.\d+(?:\.\d+)?\s+/.test(title)
    || /^(6\.\s*introduction|7\.\s*organization|8\.\s*tasks|9\.\s*case|10\.\s*reflections)/i.test(title);
}

function estimateAppendixPageCount(appendicesData?: AppendicesData): number {
  if (!appendicesData) return 1;

  let pages = 1;

  if (appendicesData.dailyJournal && appendicesData.dailyJournal.length > 0) {
    pages += appendicesData.dailyJournal.length * 2;
  }

  const appendixKeys = [
    'certParticipation',
    'primeNarrative',
    'resume',
    'grades',
    'medicalWaiver',
    'letterAcceptance',
    'dtr',
    'ratingSheet',
    'certCompletion',
  ] as const;

  appendixKeys.forEach((key) => {
    if (appendicesData[key]) pages += 1;
  });

  return Math.max(1, pages);
}

function getSectionPageMap(data: any): Record<string, number> {
  const compact = data.sectionLayout === 'compact';
  const acknowledgementLines = estimateAcknowledgementLineCount(data);
  const tocLines = estimateTOCLineCount(data);
  const acknowledgementPages = Math.max(1, Math.ceil(acknowledgementLines / ESTIMATED_PAGE_LINES));
  const tocPages = Math.max(1, Math.ceil(tocLines / ESTIMATED_PAGE_LINES));
  const narrativeStartPage = 2 + acknowledgementPages + tocPages;
  let pageCursor = compact
    ? narrativeStartPage
    : 2
      + acknowledgementPages
      + tocPages;
  let lineCursor = 0;
  const pageMap: Record<string, number> = {};
  const sectionDefinitions = [
    { key: 'introduction', value: [
      { title: 'Background of the Organization', content: data.background },
      { title: 'Vision', content: data.vision },
      { title: 'Mission', content: data.mission },
      { title: 'Objectives', content: data.objectives, isBullet: true },
      { title: 'Core Values', content: data.coreValues, isBullet: true },
      { title: 'Products and Services Offered', content: data.services },
    ], image: Boolean(data.organizationStructureImage) },
    { key: 'organization company analysis', value: [
      { title: 'Strengths', content: data.strengths, isBullet: true },
      { title: 'Weaknesses', content: data.weaknesses, isBullet: true },
      { title: 'Opportunities', content: data.opportunities, isBullet: true },
      { title: 'Threats', content: data.threats, isBullet: true },
      { title: 'Recommendations for Improvement', content: data.recommendations, isBullet: true },
    ], image: false },
    { key: 'tasks and duties', value: [
      { title: 'Assigned Tasks and Responsibilities', content: data.tasks, isBullet: true },
      { title: 'Duties and Procedures Conformed', content: data.procedures, isBullet: true },
    ], image: false },
    { key: 'case analysis', value: [
      { title: 'Issue / Problem 1', content: data.issue1 },
      { title: 'Strategy/Action Undertaken for Problem 1', content: data.issue1Action },
      { title: 'Issue / Problem 2', content: data.issue2 },
      { title: 'Strategy/Action Undertaken for Problem 2', content: data.issue2Action },
      { title: 'Lessons Learned from the Situations', content: data.lessons },
    ], image: false },
    { key: 'reflections', value: [
      { title: 'Self-Evaluation', content: data.selfEvaluation },
      { title: 'Relevancy of the Organization', content: data.relevancy },
    ], image: false },
  ];

  sectionDefinitions.forEach((section, index) => {
    const lineCount = estimateSectionLineCount(section.value, section.image);
    pageMap[section.key] = compact
      ? narrativeStartPage + Math.floor(lineCursor / ESTIMATED_PAGE_LINES)
      : pageCursor;
    addSubsectionPageMap(pageMap, String(index + 1), section.value, pageMap[section.key], compact);

    if (compact) {
      lineCursor += lineCount;
    } else {
      pageCursor += estimateSectionPageCount(section.value, section.image);
    }
  });

  if (data.schoolPlacement?.enabled) {
    const schoolDefinitions = [
      { key: 'school ojt introduction', value: [
        { title: 'School/OJT Unit', content: data.schoolPlacement.trainingOrganization },
        { title: 'Location', content: data.schoolPlacement.trainingLocation },
        { title: 'Background of the Organization', content: data.schoolPlacement.background },
        { title: 'Vision', content: data.schoolPlacement.vision },
        { title: 'Mission', content: data.schoolPlacement.mission },
        { title: 'Objectives', content: data.schoolPlacement.objectives, isBullet: true },
        { title: 'Core Values', content: data.schoolPlacement.coreValues, isBullet: true },
        { title: 'Products and Services Offered', content: data.schoolPlacement.services },
      ], image: Boolean(data.schoolPlacement.organizationStructureImage) },
      { key: 'school ojt organization analysis', value: [
        { title: 'Strengths', content: data.schoolPlacement.strengths, isBullet: true },
        { title: 'Weaknesses', content: data.schoolPlacement.weaknesses, isBullet: true },
        { title: 'Opportunities', content: data.schoolPlacement.opportunities, isBullet: true },
        { title: 'Threats', content: data.schoolPlacement.threats, isBullet: true },
        { title: 'Recommendations for Improvement', content: data.schoolPlacement.recommendations, isBullet: true },
      ], image: false },
      { key: 'school ojt tasks and duties', value: [
        { title: 'Assigned Tasks and Responsibilities', content: data.schoolPlacement.tasks, isBullet: true },
        { title: 'Duties and Procedures Conformed', content: data.schoolPlacement.procedures, isBullet: true },
      ], image: false },
      { key: 'school ojt case analysis', value: [
        { title: 'Issue / Problem 1', content: data.schoolPlacement.issue1 },
        { title: 'Strategy/Action Undertaken for Problem 1', content: data.schoolPlacement.issue1Action },
        { title: 'Issue / Problem 2', content: data.schoolPlacement.issue2 },
        { title: 'Strategy/Action Undertaken for Problem 2', content: data.schoolPlacement.issue2Action },
        { title: 'Lessons Learned from the Situations', content: data.schoolPlacement.lessons },
      ], image: false },
      { key: 'school ojt reflections', value: [
        { title: 'Self-Evaluation', content: data.schoolPlacement.selfEvaluation },
        { title: 'Relevancy of the Organization', content: data.schoolPlacement.relevancy },
      ], image: false },
    ];

    schoolDefinitions.forEach((section, index) => {
      const estimatedValue = index === 0
        ? [{ title: 'School OJT Narrative (Second Placement)', content: '' }, ...section.value]
        : section.value;
      const lineCount = estimateSectionLineCount(estimatedValue, section.image);
      const schoolSectionNumber = String(index + 6);
      pageMap[section.key] = compact
        ? narrativeStartPage + Math.floor(lineCursor / ESTIMATED_PAGE_LINES)
        : pageCursor;
      addSubsectionPageMap(pageMap, schoolSectionNumber, section.value, pageMap[section.key], compact);

      if (compact) {
        lineCursor += lineCount;
      } else {
        pageCursor += estimateSectionPageCount(estimatedValue, section.image);
      }
    });

    pageMap['school ojt narrative'] = pageMap['school ojt introduction'];
  }

  pageMap.appendices = compact
    ? narrativeStartPage + Math.ceil(lineCursor / ESTIMATED_PAGE_LINES)
    : pageCursor;
  return pageMap;
}

function buildTableOfContentsPage(data: any, compact = false, lineSpacing = 240): Paragraph[] {
  const rawContent = data.tableOfContents || '';
  const sectionPageMap = getSectionPageMap(data);

  const fallbackEntries = [
    { title: '1. Introduction', page: String(sectionPageMap.introduction || 4) },
    { title: '1.1 Background of the Organization', page: String(sectionPageMap.introduction || 4) },
    { title: '1.2 Vision', page: String(sectionPageMap.introduction || 4) },
    { title: '1.3 Mission', page: String(sectionPageMap.introduction || 4) },
    { title: '1.4 Objectives', page: String(sectionPageMap.introduction || 4) },
    { title: '1.5 Core Values', page: String(sectionPageMap.introduction || 4) },
    { title: '1.6 Products and Services Offered', page: String(sectionPageMap.introduction || 4) },
    { title: '2. Organization / Company Analysis', page: String(sectionPageMap['organization company analysis'] || 5) },
    { title: '2.1 Strengths', page: String(sectionPageMap['organization company analysis'] || 5) },
    { title: '2.2 Weaknesses', page: String(sectionPageMap['organization company analysis'] || 5) },
    { title: '2.3 Opportunities', page: String(sectionPageMap['organization company analysis'] || 5) },
    { title: '2.4 Threats', page: String(sectionPageMap['organization company analysis'] || 5) },
    { title: '2.5 Recommendations for Improvement', page: String(sectionPageMap['organization company analysis'] || 5) },
    { title: '3. Tasks and Duties', page: String(sectionPageMap['tasks and duties'] || 6) },
    { title: '3.1 Assigned Tasks and Responsibilities', page: String(sectionPageMap['tasks and duties'] || 6) },
    { title: '3.2 Duties and Procedures Conformed', page: String(sectionPageMap['tasks and duties'] || 6) },
    { title: '4. Case Analysis', page: String(sectionPageMap['case analysis'] || 7) },
    { title: '4.1 Issue / Problem 1', page: String(sectionPageMap['case analysis'] || 7) },
    { title: '4.2 Strategy/Action Undertaken for Problem 1', page: String(sectionPageMap['case analysis'] || 7) },
    { title: '4.3 Issue / Problem 2', page: String(sectionPageMap['case analysis'] || 7) },
    { title: '4.4 Strategy/Action Undertaken for Problem 2', page: String(sectionPageMap['case analysis'] || 7) },
    { title: '4.5 Lessons Learned from the Situations', page: String(sectionPageMap['case analysis'] || 7) },
    { title: '5. Reflections', page: String(sectionPageMap.reflections || 8) },
    { title: '5.1 Self-Evaluation', page: String(sectionPageMap.reflections || 8) },
    { title: '5.2 Relevancy of the Organization', page: String(sectionPageMap.reflections || 8) },
    { title: '6. Appendices', page: String(sectionPageMap.appendices || 9) },
  ];

  const parsedEntries = ensureArray(rawContent).reduce<{ title: string; page: string }[]>((list, line) => {
    const trimmed = line.trim();
    if (!trimmed) return list;

    const match = trimmed.match(/^(.*?)(?:\s+)(\d+)$/) || trimmed.match(/^(.*)$/);
    const extractedTitle = (match?.[1] || trimmed)
      .replace(/\s*[:.-]+\s*$/, '')
      .trim();
    const extractedPage = match?.[2] || '';

    if (extractedTitle) {
      list.push({
        title: extractedTitle,
        page: extractedPage,
      });
    }

    return list;
  }, []);

  let entries = parsedEntries.length > 0 ? parsedEntries : fallbackEntries;
  if (data.schoolPlacement?.enabled) {
    entries = entries.filter((entry) => !isSchoolTOCLine(entry.title));
    const schoolTOCEntries = [
      { title: '6. Introduction', page: String(sectionPageMap['school ojt introduction'] || 9) },
      { title: '6.1 Background of the Organization', page: String(sectionPageMap['school ojt introduction'] || 9) },
      { title: '6.2 Vision', page: String(sectionPageMap['school ojt introduction'] || 9) },
      { title: '6.3 Mission', page: String(sectionPageMap['school ojt introduction'] || 9) },
      { title: '6.4 Objectives', page: String(sectionPageMap['school ojt introduction'] || 9) },
      { title: '6.5 Core Values', page: String(sectionPageMap['school ojt introduction'] || 9) },
      { title: '6.6 Products and Services Offered', page: String(sectionPageMap['school ojt introduction'] || 9) },
      { title: '7. Organization / Company Analysis', page: String(sectionPageMap['school ojt organization analysis'] || 9) },
      { title: '7.1 Strengths', page: String(sectionPageMap['school ojt organization analysis'] || 9) },
      { title: '7.2 Weaknesses', page: String(sectionPageMap['school ojt organization analysis'] || 9) },
      { title: '7.3 Opportunities', page: String(sectionPageMap['school ojt organization analysis'] || 9) },
      { title: '7.4 Threats', page: String(sectionPageMap['school ojt organization analysis'] || 9) },
      { title: '7.5 Recommendations for Improvement', page: String(sectionPageMap['school ojt organization analysis'] || 9) },
      { title: '8. Tasks and Duties', page: String(sectionPageMap['school ojt tasks and duties'] || 9) },
      { title: '8.1 Assigned Tasks and Responsibilities', page: String(sectionPageMap['school ojt tasks and duties'] || 9) },
      { title: '8.2 Duties and Procedures Conformed', page: String(sectionPageMap['school ojt tasks and duties'] || 9) },
      { title: '9. Case Analysis', page: String(sectionPageMap['school ojt case analysis'] || 9) },
      { title: '9.1 Issue / Problem 1', page: String(sectionPageMap['school ojt case analysis'] || 9) },
      { title: '9.2 Strategy/Action Undertaken for Problem 1', page: String(sectionPageMap['school ojt case analysis'] || 9) },
      { title: '9.3 Issue / Problem 2', page: String(sectionPageMap['school ojt case analysis'] || 9) },
      { title: '9.4 Strategy/Action Undertaken for Problem 2', page: String(sectionPageMap['school ojt case analysis'] || 9) },
      { title: '9.5 Lessons Learned from the Situations', page: String(sectionPageMap['school ojt case analysis'] || 9) },
      { title: '10. Reflections', page: String(sectionPageMap['school ojt reflections'] || 9) },
      { title: '10.1 Self-Evaluation', page: String(sectionPageMap['school ojt reflections'] || 9) },
      { title: '10.2 Relevancy of the Organization', page: String(sectionPageMap['school ojt reflections'] || 9) },
    ];
    const appendicesIndex = entries.findIndex((entry) => normalizeTOCKey(entry.title) === 'appendices');
    if (appendicesIndex >= 0) entries[appendicesIndex] = { ...entries[appendicesIndex], title: '11. Appendices' };
    entries.splice(appendicesIndex >= 0 ? appendicesIndex : entries.length, 0, ...schoolTOCEntries);
  } else {
    entries = entries.filter((entry) => !isSchoolTOCLine(entry.title));
  }

  const paragraphs: Paragraph[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 240, line: lineSpacing, lineRule: 'auto' },
      children: [new TextRun({ text: 'Table of Contents', bold: true, size: 28, font: 'Times New Roman' })],
    }),
  ];

  entries.forEach((entry) => {
    const titleText = entry.title;
    const normalizedTitle = normalizeTOCKey(titleText);
    const sectionNumber = titleText.match(/^\s*(\d+)(?:\.\d+)*[.)]?\s+/)?.[1];
    const numberedSectionKeys: Record<string, string> = {
      '1': 'introduction',
      '2': 'organization company analysis',
      '3': 'tasks and duties',
      '4': 'case analysis',
      '5': 'reflections',
      '6': 'appendices',
    };
    const numberedPage = sectionNumber ? sectionPageMap[numberedSectionKeys[sectionNumber]] : undefined;
    const subsectionMatch = titleText.match(/^\s*(\d+)\.(\d+)/);
    const subsectionPage = subsectionMatch
      ? sectionPageMap[`${subsectionMatch[1]}.${subsectionMatch[2]}`]
      : undefined;
    const schoolMainNumber = titleText.match(/^\s*(\d+)(?:\.\d+)*[.)]?\s+/)?.[1];
    const schoolSectionKeys: Record<string, string> = {
      '6': 'school ojt introduction',
      '7': 'school ojt organization analysis',
      '8': 'school ojt tasks and duties',
      '9': 'school ojt case analysis',
      '10': 'school ojt reflections',
    };
    const schoolPlacementPage = data.schoolPlacement?.enabled && !normalizedTitle.includes('appendices')
      ? sectionPageMap[schoolSectionKeys[schoolMainNumber || '']]
      : undefined;
    const matchedPage = Object.entries(sectionPageMap).find(([key]) => {
      const normalizedKey = normalizeTOCKey(key);
      return normalizedKey === normalizedTitle
        || normalizedTitle.startsWith(normalizedKey)
        || normalizedKey.startsWith(normalizedTitle);
    })?.[1];
    const mappedPage = subsectionPage ?? schoolPlacementPage ?? matchedPage ?? numberedPage;
    const resolvedPage = mappedPage ? String(mappedPage) : entry.page;
    const isSubsection = /^\d+\.\d+\s+/.test(titleText);
    const showPageNumber = (compact || !isSubsection) && Boolean(resolvedPage);
    const fillLength = Math.max(1, showPageNumber ? 90 - titleText.length : 110 - titleText.length);
    const filler = '.'.repeat(fillLength);
    const indent = isSubsection ? 500 : 180;

    const pageText = showPageNumber ? [
      new TextRun({ text: '\t', size: 22, font: 'Times New Roman' }),
      new TextRun({ text: resolvedPage, size: 22, font: 'Times New Roman', bold: false }),
    ] : [];

    paragraphs.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        indent: { left: indent },
        tabStops: [{ type: 'right', position: 9000 }],
        spacing: { before: 0, after: 4, line: lineSpacing, lineRule: 'auto' },
        children: [
          new TextRun({ text: `${titleText}${filler}`, size: 22, font: 'Times New Roman' }),
          ...pageText,
        ],
      })
    );
  });

  return paragraphs;
}

function buildSectionPage(title: string, sections: SectionData[], compact = false, lineSpacing = 240, imageData?: string): Paragraph[] {
  const children: any[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
      children: [new TextRun({ text: title, bold: true, size: 24, font: 'Times New Roman' })],
    }),
  ];

  sections.forEach((section) => {
    const text = typeof section.content === 'string' ? section.content : Array.isArray(section.content) ? section.content.join('\n') : String(section.content ?? '');
    const contentLines = ensureArray(text);

    if (section.title) {
      children.push(
        new Paragraph({
          spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
          children: [new TextRun({ text: `${section.title}:`, bold: true, size: 24, font: 'Times New Roman' })],
        })
      );
    }

    if (section.isBullet) {
      contentLines.forEach((line) => {
        children.push(
          new Paragraph({
            bullet: { level: 0 },
            indent: { left: 720 },
            spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
            children: [new TextRun({ text: line, size: 24, font: 'Times New Roman' })],
          })
        );
      });
    } else if (contentLines.length > 0) {
      const combinedText = contentLines.join(' ');
      children.push(
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
          children: [new TextRun({ text: combinedText, size: 24, font: 'Times New Roman' })],
        })
      );
    }
  });

  const imageBuffer = parseBase64Image(imageData || '');
  if (imageBuffer) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 60 },
        children: [new TextRun({ text: 'Organization Structure', bold: true, size: 24, font: 'Times New Roman' })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 120 },
        children: [new ImageRun({ type: 'png', data: imageBuffer, transformation: { width: 520, height: 360 } })],
      })
    );
  }

  return children;
}

function buildSchoolPlacementPages(placement: any, compact = false, lineSpacing = 240): Paragraph[] {
  const pages: Paragraph[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
      children: [new TextRun({ text: 'SCHOOL OJT NARRATIVE (SECOND PLACEMENT)', bold: true, size: 24, font: 'Times New Roman' })],
    }),
  ];
  const sections = [
    {
      title: 'SCHOOL OJT - INTRODUCTION',
      content: [
        { title: 'School/OJT Unit', content: placement.trainingOrganization },
        { title: 'Location', content: placement.trainingLocation },
        { title: 'Background of the Organization', content: placement.background },
        { title: 'Vision', content: placement.vision },
        { title: 'Mission', content: placement.mission },
        { title: 'Objectives', content: placement.objectives, isBullet: true },
        { title: 'Core Values', content: placement.coreValues, isBullet: true },
        { title: 'Products and Services Offered', content: placement.services },
      ],
      image: placement.organizationStructureImage,
    },
    {
      title: 'SCHOOL OJT - ORGANIZATION ANALYSIS',
      content: [
        { title: 'Strengths', content: placement.strengths, isBullet: true },
        { title: 'Weaknesses', content: placement.weaknesses, isBullet: true },
        { title: 'Opportunities', content: placement.opportunities, isBullet: true },
        { title: 'Threats', content: placement.threats, isBullet: true },
        { title: 'Recommendations for Improvement', content: placement.recommendations, isBullet: true },
      ],
    },
    {
      title: 'SCHOOL OJT - TASKS AND DUTIES',
      content: [
        { title: 'Assigned Tasks and Responsibilities', content: placement.tasks, isBullet: true },
        { title: 'Duties and Procedures Conformed', content: placement.procedures, isBullet: true },
      ],
    },
    {
      title: 'SCHOOL OJT - CASE ANALYSIS',
      content: [
        { title: 'Issue / Problem 1', content: placement.issue1 },
        { title: 'Strategy/Action Undertaken for Problem 1', content: placement.issue1Action },
        { title: 'Issue / Problem 2', content: placement.issue2 },
        { title: 'Strategy/Action Undertaken for Problem 2', content: placement.issue2Action },
        { title: 'Lessons Learned from the Situations', content: placement.lessons },
      ],
    },
    {
      title: 'SCHOOL OJT - REFLECTIONS',
      content: [
        { title: 'Self-Evaluation', content: placement.selfEvaluation },
        { title: 'Relevancy of the Organization', content: placement.relevancy },
      ],
    },
  ];

  sections.forEach((section, index) => {
    if (index > 0 && !compact) pages.push(new Paragraph({ pageBreakBefore: true }));
    pages.push(...buildSectionPage(section.title, section.content, compact, lineSpacing, section.image));
  });

  return pages;
}

function buildImageGridTable(images: (string | AppendixImage)[]) {
  const rows: TableRow[] = [];
  const invisibleBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const tableBorders = { top: invisibleBorder, bottom: invisibleBorder, left: invisibleBorder, right: invisibleBorder, insideHorizontal: invisibleBorder, insideVertical: invisibleBorder };

  for (let i = 0; i < images.length; i += 2) {
    const img1 = images[i];
    const img2 = images[i + 1];

    const cell1Children: any[] = [];
    const cell2Children: any[] = [];

    if (img1) {
      const isObj1 = typeof img1 === 'object' && img1 !== null;
      const base64_1 = isObj1 ? (img1 as AppendixImage).base64 : (img1 as string);
      const detail_1 = isObj1 ? (img1 as AppendixImage).detail : '';
      const buf1 = parseBase64Image(base64_1);

      if (buf1) {
        cell1Children.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 60 },
            children: [
              new ImageRun({
                type: 'png',
                data: buf1,
                transformation: { width: 280, height: 210 },
              }),
            ],
          })
        );
        if (detail_1) {
          cell1Children.push(
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 120 },
              children: [new TextRun({ text: detail_1, size: 24, font: 'Times New Roman' })],
            })
          );
        }
      }
    }

    if (img2) {
      const isObj2 = typeof img2 === 'object' && img2 !== null;
      const base64_2 = isObj2 ? (img2 as AppendixImage).base64 : (img2 as string);
      const detail_2 = isObj2 ? (img2 as AppendixImage).detail : '';
      const buf2 = parseBase64Image(base64_2);

      if (buf2) {
        cell2Children.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 60 },
            children: [
              new ImageRun({
                type: 'png',
                data: buf2,
                transformation: { width: 280, height: 210 },
              }),
            ],
          })
        );
        if (detail_2) {
          cell2Children.push(
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 120 },
              children: [new TextRun({ text: detail_2, size: 24, font: 'Times New Roman' })],
            })
          );
        }
      }
    }

    rows.push(
      new TableRow({
        children: [
          new TableCell({ children: cell1Children, width: { size: 50, type: WidthType.PERCENTAGE } }),
          new TableCell({ children: cell2Children, width: { size: 50, type: WidthType.PERCENTAGE } }),
        ],
      })
    );
  }

  return new Table({
    rows,
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: tableBorders,
  });
}

function buildFeaturedImageLayout(imgData: string | AppendixImage) {
  const children: any[] = [];
  const isObj = typeof imgData === 'object' && imgData !== null;
  const base64 = isObj ? (imgData as AppendixImage).base64 : (imgData as string);
  const detail = isObj ? (imgData as AppendixImage).detail : '';
  const buf = parseBase64Image(base64);

  if (buf) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 100 },
        children: [
          new ImageRun({
            type: 'png',
            data: buf,
            transformation: { width: 420, height: 315 },
          }),
        ],
      })
    );

    if (detail) {
      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [new TextRun({ text: detail, size: 24, font: 'Times New Roman' })],
        })
      );
    }
  }

  return children;
}

function buildReportFooterBlock(footer?: ReportFooterData): any[] {
  const preparedByLabel = footer?.preparedByLabel?.trim() || 'Prepared By:';
  const preparedByName = footer?.preparedByName?.trim() || 'HAYNA G. DAUD';
  const checkedByLabel = footer?.checkedByLabel?.trim() || 'Checked By:';
  const checkedByName = footer?.checkedByName?.trim() || 'ANDRES S. TAPALES JR.';
  const checkedByRole = footer?.checkedByRole?.trim() || 'Supervising Statistical Specialist';
  const dateLabel = footer?.dateLabel?.trim() || 'Date:';
  const dateValue = footer?.dateValue?.trim() || 'August 15, 2026';
  const officeInCharge = footer?.officeInCharge?.trim() || 'Office-in-Charge';

  const footerTable = new Table({
    width: { size: 94, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      insideHorizontal: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      insideVertical: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 10, bottom: 10, left: 10, right: 10 },
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                spacing: { after: 12, line: 276, lineRule: 'auto' },
                children: [new TextRun({ text: preparedByLabel, bold: true, size: 24, font: 'Times New Roman' })],
              })
            ],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 10, bottom: 10, left: 10, right: 10 },
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                spacing: { after: 12, line: 276, lineRule: 'auto' },
                children: [new TextRun({ text: checkedByLabel, bold: true, size: 24, font: 'Times New Roman' })],
              })
            ],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 8, bottom: 8, left: 10, right: 10 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 6, after: 10, line: 276, lineRule: 'auto' },
                children: [new TextRun({ text: preparedByName, bold: true, size: 24, font: 'Times New Roman' })],
              })
            ],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 8, bottom: 8, left: 10, right: 10 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 6, after: 10, line: 276, lineRule: 'auto' },
                children: [new TextRun({ text: checkedByName, bold: true, size: 24, font: 'Times New Roman' })],
              })
            ],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 10, bottom: 10, left: 10, right: 10 },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 8, after: 8, line: 276, lineRule: 'auto' }, children: [new TextRun({ text: '', size: 24, font: 'Times New Roman' })] })],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 8, bottom: 8, left: 10, right: 10 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 8, after: 10, line: 276, lineRule: 'auto' },
                children: [new TextRun({ text: checkedByRole, size: 24, font: 'Times New Roman' })],
              })
            ],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 8, bottom: 10, left: 10, right: 10 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 8, after: 8, line: 276, lineRule: 'auto' },
                children: [
                  new TextRun({ text: `${dateLabel} `, bold: true, size: 24, font: 'Times New Roman' }),
                  new TextRun({ text: dateValue, bold: true, size: 24, font: 'Times New Roman' }),
                ],
              })
            ],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 8, bottom: 10, left: 10, right: 10 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 8, after: 8, line: 276, lineRule: 'auto' },
                children: [new TextRun({ text: officeInCharge, size: 24, font: 'Times New Roman' })],
              })
            ],
          }),
        ],
      }),
    ],
  });

  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0, line: 276, lineRule: 'auto' },
      children: [footerTable],
    }),
  ];
}

function buildDailyJournalAppendixPage(appendicesData: AppendicesData) {
  const children: any[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 360 },
      children: [new TextRun({ text: 'APPENDICES', bold: true, size: 24, font: 'Times New Roman' })],
    }),
  ];

  if (!appendicesData || !appendicesData.dailyJournal || !Array.isArray(appendicesData.dailyJournal)) return children;

  const layout = appendicesData.dailyJournalLayout === 'no-photo'
    ? 'no-photo'
    : appendicesData.dailyJournalLayout === 'report' ? 'report' : 'current';

  if (layout === 'no-photo') {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 180, after: 180 },
        children: [new TextRun({ text: 'DAILY JOURNAL', bold: true, size: 24, font: 'Times New Roman' })],
      })
    );

    appendicesData.dailyJournal.forEach((weekData) => {
      const totalHours = weekData.activities.reduce((total, activity) => total + (Number(activity.hours) || 0), 0);
      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 120, after: 60 },
          keepNext: true,
          children: [new TextRun({ text: `WEEK ${weekData.weekNumber}`, bold: true, size: 24, font: 'Times New Roman' })],
        }),
        buildCurrentWeeklyTable(weekData.activities, totalHours),
      );
    });

    return children;
  }

  if (layout === 'current') {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 180, after: 90 },
        children: [new TextRun({ text: 'WEEKLY WORK ACTIVITIES', bold: true, size: 24, font: 'Times New Roman' })],
      })
    );
  }

  const narrativeLineSpacing = 360;

  appendicesData.dailyJournal.forEach((weekData: DailyJournalWeek, index: number) => {
    if (layout === 'report') {
      // Restore the original daily layout behavior, which keeps documentation inside the table.
      const reportImages = weekData.reportLayoutImages || weekData.images || [];
      const weekParagraph = new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 120 },
        pageBreakBefore: index > 0,
        children: [new TextRun({ text: `WEEK ${weekData.weekNumber}`, bold: true, size: 24, font: 'Times New Roman' })],
      });
      children.push(
        weekParagraph,
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 60, after: 120 },
          children: [new TextRun({ text: 'WEEKLY ACCOMPLISHMENT REPORT', bold: true, size: 24, font: 'Times New Roman' })],
        }),
        buildDailyReportTable(weekData.activities, reportImages),
      );

      return;
    }

    // Use current layout images, fall back to shared images for backwards compatibility
    const currentImages = weekData.currentLayoutImages || weekData.images || [];
    const currentWeekParagraph = new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 80 },
      pageBreakBefore: index > 0,
      children: [new TextRun({ text: `WEEK ${weekData.weekNumber}`, bold: true, size: 24, font: 'Times New Roman' })],
    });
    children.push(
      currentWeekParagraph,
      buildCurrentWeeklyTable(weekData.activities, weekData.totalHours),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 120 },
        children: [new TextRun({ text: `Table ${weekData.weekNumber}: Week ${weekData.weekNumber}`, bold: true, size: 24, font: 'Times New Roman' })],
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { before: 0, after: 0, line: narrativeLineSpacing, lineRule: 'auto' },
        children: [new TextRun({ text: weekData.narrative || 'No narrative provided.', size: 24, font: 'Times New Roman' })],
      })
    );

    if (currentImages.length > 0) {
      children.push(buildImageGridTable(currentImages));
    }
  });

  return children;
}

function buildOtherAppendixPages(appendicesData: AppendicesData) {
  const children: any[] = [];
  if (!appendicesData) return children;

  const appendixList = [
    { title: 'Certificate of Participation (PRIME Seminar)', key: 'certParticipation' as const },
    { title: 'One page narrative report of PRIME', key: 'primeNarrative' as const, isText: true },
    { title: 'Resume and Application Letter', key: 'resume' as const },
    { title: 'Evaluation of Grades & Validated Enrolment Form', key: 'grades' as const },
    { title: 'Medical Certificate & Parent\'s Waiver', key: 'medicalWaiver' as const },
    { title: 'Letter of Acceptance', key: 'letterAcceptance' as const },
    { title: 'Daily Time Record (DTR) - hours rendered', key: 'dtr' as const },
    { title: 'Student’s Intern Rating Sheet & Performance Evaluation', key: 'ratingSheet' as const },
    { title: 'Certificate of Completion & Memorandum of Agreement', key: 'certCompletion' as const },
  ];

  appendixList.forEach(app => {
    const data = appendicesData[app.key];
    if (!data) return;

    if (children.length > 0) {
      children.push(new Paragraph({ pageBreakBefore: true }));
    }
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 240, after: 240 },
        children: [new TextRun({ text: app.title, bold: true, size: 24, font: 'Times New Roman' })],
      })
    );

    if (app.isText) {
      if (app.key === 'primeNarrative') {
        const primeImages = (appendicesData.primeReportImages || []).slice(0, 2);
        if (primeImages.length > 0) {
          children.push(buildImageGridTable(primeImages));
        }
        children.push(
          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { before: 0, after: 0, line: 360, lineRule: 'auto' },
            children: [new TextRun({ text: String(data), size: 24, font: 'Times New Roman' })],
          })
        );
        return;
      }

      children.push(
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun({ text: data, size: 24, font: 'Times New Roman' })],
        })
      );
    } else {
      children.push(...buildFeaturedImageLayout(data));
    }
  });

  return children;
}

function buildAppendicesPage(appendicesData: AppendicesData) {
  return [
    ...buildDailyJournalAppendixPage(appendicesData),
    ...buildOtherAppendixPages(appendicesData),
  ];
}

function buildCurrentWeeklyTable(activities: DailyActivity[], totalHours: number | string) {
  const TABLE_TEXT_SIZE = 24;
  const tableRows = [
    new TableRow({
      cantSplit: true,
      children: [
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 28, after: 28 }, children: [new TextRun({ text: 'DAY', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 16, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 28, after: 28 }, children: [new TextRun({ text: 'DATE', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 18, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 28, after: 28 }, children: [new TextRun({ text: 'ACCOMPLISHMENT', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 52, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 28, after: 28 }, children: [new TextRun({ text: 'NO. OF WORKING HOURS', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 14, type: WidthType.PERCENTAGE } }),
      ],
    }),
  ];

  const rows = activities && activities.length > 0 ? activities : [{ day: 'Day 1', date: '', accomplishment: '', hours: '' }];

  rows.forEach((act) => {
    tableRows.push(
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({
            width: { size: 16, type: WidthType.PERCENTAGE },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 18, after: 18 }, children: [new TextRun({ text: act.day || '', size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })],
            verticalAlign: 'center',
          }),
          new TableCell({
            width: { size: 18, type: WidthType.PERCENTAGE },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 18, after: 18 }, children: [new TextRun({ text: act.date || '', size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })],
            verticalAlign: 'center',
          }),
          new TableCell({
            width: { size: 52, type: WidthType.PERCENTAGE },
            children: (() => {
              const accomplishmentText = act.accomplishment || 'No accomplishment added.';
              const lines = accomplishmentText.split('\n').filter(line => line.trim());
              if (lines.length === 0) {
                return [new Paragraph({ alignment: AlignmentType.LEFT, spacing: { before: 18, after: 18 }, children: [new TextRun({ text: 'No accomplishment added.', size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })];
              }
              return lines.map((line, idx) =>
                new Paragraph({
                  alignment: AlignmentType.LEFT,
                  spacing: { before: 14, after: idx === lines.length - 1 ? 14 : 8 },
                  children: [new TextRun({ text: line.trim(), size: TABLE_TEXT_SIZE, font: 'Times New Roman' })]
                })
              );
            })(),
          }),
          new TableCell({
            width: { size: 14, type: WidthType.PERCENTAGE },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 18, after: 18 }, children: [new TextRun({ text: String(act.hours ?? ''), size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })],
            verticalAlign: 'center',
          }),
        ],
      })
    );
  });

  tableRows.push(
    new TableRow({
      cantSplit: true,
      children: [
        new TableCell({ 
          columnSpan: 3, 
          children: [new Paragraph({ alignment: AlignmentType.LEFT, spacing: { before: 24, after: 24 }, children: [new TextRun({ text: 'TOTAL NUMBER OF HOURS', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })] 
        }),
        new TableCell({ 
          children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 24, after: 24 }, children: [new TextRun({ text: String(totalHours), bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })] 
        }),
      ],
    })
  );

  return new Table({
    rows: tableRows,
    alignment: AlignmentType.CENTER,
    width: { size: 96, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    columnWidths: [1600, 1800, 5200, 1400],
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
    },
  });
}

function buildDailyReportTable(activities: DailyActivity[], images: (string | AppendixImage)[] = []) {
  const TABLE_TEXT_SIZE = 24;
  const tableRows = [
    new TableRow({
      cantSplit: true,
      children: [
        new TableCell({ margins: { top: 18, bottom: 18, left: 24, right: 24 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: 'DATE', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 18, type: WidthType.PERCENTAGE } }),
        new TableCell({ margins: { top: 18, bottom: 18, left: 24, right: 24 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: 'ACCOMPLISHMENT', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 34, type: WidthType.PERCENTAGE } }),
        new TableCell({ margins: { top: 18, bottom: 18, left: 24, right: 24 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: 'HOURS', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 10, type: WidthType.PERCENTAGE } }),
        new TableCell({ margins: { top: 18, bottom: 18, left: 24, right: 24 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: 'DOCUMENTATION', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 38, type: WidthType.PERCENTAGE } }),
      ],
    }),
  ];

  const rows = activities && activities.length > 0 ? activities : [{ day: 'Day 1', date: '', accomplishment: '', hours: '' }];

  rows.forEach((act, index) => {
    const imageData = images[index] || null;
    const imageCellChildren: any[] = [];

    if (imageData) {
      const isObj = typeof imageData === 'object' && imageData !== null;
      const base64 = isObj ? (imageData as AppendixImage).base64 : (imageData as string);
      const detail = isObj ? (imageData as AppendixImage).detail : '';
      const buf = parseBase64Image(base64);

      if (buf) {
        imageCellChildren.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 0 },
            children: [new ImageRun({ type: 'png', data: buf, transformation: { width: 260, height: 170 } })],
          })
        );

        if (detail) {
          imageCellChildren.push(
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 0, after: 0 },
              children: [new TextRun({ text: detail, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })],
            })
          );
        }
      }
    }

    tableRows.push(
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({ margins: { top: 12, bottom: 12, left: 18, right: 18 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: act.date || act.day || '', size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], verticalAlign: 'center' }),
          new TableCell({
            margins: { top: 12, bottom: 12, left: 18, right: 18 },
            children: (() => {
              const accomplishmentText = act.accomplishment || 'No accomplishment added.';
              const lines = accomplishmentText.split('\n').filter(line => line.trim());
              if (lines.length === 0) {
                return [new Paragraph({ alignment: AlignmentType.LEFT, spacing: { before: 0, after: 0 }, bullet: { level: 0 }, children: [new TextRun({ text: 'No accomplishment added.', size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })];
              }
              return lines.map((line, idx) =>
                new Paragraph({
                  alignment: AlignmentType.LEFT,
                  spacing: { before: 0, after: idx === lines.length - 1 ? 0 : 0 },
                  bullet: { level: 0 },
                  children: [new TextRun({ text: line.trim(), size: TABLE_TEXT_SIZE, font: 'Times New Roman' })]
                })
              );
            })()
          }),
          new TableCell({ margins: { top: 12, bottom: 12, left: 18, right: 18 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: String(act.hours ?? ''), size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], verticalAlign: 'center' }),
          new TableCell({ margins: { top: 12, bottom: 12, left: 18, right: 18 }, children: imageCellChildren.length > 0 ? imageCellChildren : [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '', size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })] }),
        ],
      })
    );
  });

  return new Table({
    rows: tableRows,
    alignment: AlignmentType.CENTER,
    width: { size: 94, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    columnWidths: [1700, 3300, 1000, 4000],
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
    },
  });
}

function buildWeeklyReportTable(activities: DailyActivity[], images: (string | AppendixImage)[] = []) {
  const TABLE_TEXT_SIZE = 24;
  const tableRows = [
    new TableRow({
      cantSplit: true,
      children: [
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'DAY', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 16, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'DATE', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 18, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'ACCOMPLISHMENT', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 52, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'NO. OF WORKING HOURS', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 14, type: WidthType.PERCENTAGE } }),
      ],
    }),
  ];

  const rows = activities && activities.length > 0 ? activities : [{ day: 'Day 1', date: '', accomplishment: '', hours: '' }];

  rows.forEach((act) => {
    tableRows.push(
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: act.day || '', size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], verticalAlign: 'center' }),
          new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: act.date || '', size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], verticalAlign: 'center' }),
          new TableCell({ 
            children: (() => {
              const accomplishmentText = act.accomplishment || 'No accomplishment added.';
              const lines = accomplishmentText.split('\n').filter(line => line.trim());
              if (lines.length === 0) {
                return [new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 40 }, children: [new TextRun({ text: 'No accomplishment added.', size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })];
              }
              return lines.map((line, idx) => 
                new Paragraph({ 
                  alignment: AlignmentType.LEFT, 
                  spacing: { after: idx === lines.length - 1 ? 40 : 20 },
                  children: [new TextRun({ text: line.trim(), size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] 
                })
              );
            })()
          }),
          new TableCell({
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(act.hours ?? ''), size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })],
            verticalAlign: 'center',
          }),
        ],
      })
    );
  });

  const totalCellValue = rows.reduce((sum, row) => sum + Number(row.hours || 0), 0);

  tableRows.push(
    new TableRow({
      cantSplit: true,
      children: [
        new TableCell({
          columnSpan: 3,
          children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: 'TOTAL NUMBER OF HOURS', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })],
        }),
        new TableCell({
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(totalCellValue), bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })],
        }),
      ],
    })
  );

  return new Table({
    rows: tableRows,
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
    },
  });
}
