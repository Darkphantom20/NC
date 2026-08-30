import fs from 'fs';
import path from 'path';
import { createClient } from '@supabase/supabase-js';
import { Document, Packer, Paragraph, TextRun, AlignmentType, ImageRun, Header, Footer, Table, TableRow, TableCell, WidthType, BorderStyle } from 'docx';

export const runtime = 'nodejs';

const OUTPUT_FILE_NAME = process.env.OUTPUT_FILE_NAME || 'JRMSU_Narrative_Report.docx';
const HEADER_IMAGE_PATH = path.join(process.cwd(), 'public', 'Picture1.png');
const HEADER_IMAGE = fs.existsSync(HEADER_IMAGE_PATH) ? fs.readFileSync(HEADER_IMAGE_PATH) : null;

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
  images?: (string | AppendixImage)[];
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
  dailyJournalLayout?: 'current' | 'report';
  reportFooter?: ReportFooterData;
  dailyJournal?: DailyJournalWeek[];
  certParticipation?: string;
  primeNarrative?: string;
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

function buildHeaderImage() {
  if (!HEADER_IMAGE) return [];
  return [
    new ImageRun({
      data: HEADER_IMAGE,
      type: 'png',
      transformation: { width: 624, height: 125 },
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

    if (hasSupabaseConfig) {
      const supabase = createClient(supabaseUrl!, supabaseKey!);
      const payload = {
        student_name: data.studentName || null,
        appendices: JSON.stringify(data.appendices) || null,
      };
      await supabase.from('reports').insert([payload]);
    }

    const reportFooterChildren = data.appendices?.dailyJournalLayout === 'report'
      ? buildReportFooterBlock(data.appendices?.reportFooter)
      : [];

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
          ...buildAcknowledgementPage(data),
          new Paragraph({ pageBreakBefore: true }),
          ...buildSectionPage('INTRODUCTION', [
            { title: 'Background of the Organization', content: data.background },
            { title: 'Vision', content: data.vision },
            { title: 'Mission', content: data.mission },
            { title: 'Objectives', content: data.objectives, isBullet: true },
            { title: 'Core Values', content: data.coreValues, isBullet: true },
            { title: 'Products and Services Offered', content: data.services },
          ]),
          new Paragraph({ pageBreakBefore: true }),
          ...buildSectionPage('ORGANIZATION / COMPANY ANALYSIS', [
            { title: 'Strengths', content: data.strengths, isBullet: true },
            { title: 'Weaknesses', content: data.weaknesses, isBullet: true },
            { title: 'Opportunities', content: data.opportunities, isBullet: true },
            { title: 'Threats', content: data.threats, isBullet: true },
            { title: 'Recommendations for Improvement', content: data.recommendations, isBullet: true },
          ]),
          new Paragraph({ pageBreakBefore: true }),
          ...buildSectionPage('TASKS AND DUTIES', [
            { title: 'Assigned Tasks and Responsibilities', content: data.tasks, isBullet: true },
            { title: 'Duties and Procedures Conformed', content: data.procedures, isBullet: true },
          ]),
          new Paragraph({ pageBreakBefore: true }),
          ...buildSectionPage('CASE ANALYSIS', [
            { title: 'Issue / Problem 1', content: data.issue1 },
            { title: 'Strategy/Action Undertaken for Problem 1', content: data.issue1Action },
            { title: 'Issue / Problem 2', content: data.issue2 },
            { title: 'Strategy/Action Undertaken for Problem 2', content: data.issue2Action },
            { title: 'Lessons Learned from the Situations', content: data.lessons },
          ]),
          new Paragraph({ pageBreakBefore: true }),
          ...buildSectionPage('REFLECTIONS', [
            { title: 'Self-Evaluation', content: data.selfEvaluation },
            { title: 'Relevancy of the Organization', content: data.relevancy },
          ]),
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
        children: dailyJournalContent,
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
        children: remainingAppendixContent,
      });
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

function buildCoverPage(data: any): Paragraph[] {
  const org = data.trainingOrganization || '';
  const location = data.trainingLocation || '';
  const faculty = data.collegeFaculty || '';
  const degree = data.degreeProgram || '';
  const student = data.studentName || '';
  const studentDegree = data.studentDegree || '';
  const adviserName = data.submittedToName || '';
  const adviserTitle = data.submittedToTitle || '';

  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 1440, after: 360 },
      children: [new TextRun({ text: 'A Narrative Report on the', size: 26, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 360 },
      children: [new TextRun({ text: 'On-the-Job Training conducted at', size: 26, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 180 },
      children: [new TextRun({ text: org, size: 30, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 1440 },
      children: [new TextRun({ text: location, size: 26, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 360 },
      children: [new TextRun({ text: 'Presented to the faculty of', size: 26, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 1440 },
      children: [new TextRun({ text: faculty, size: 30, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 360 },
      children: [new TextRun({ text: 'In partial fulfillment of the requirements for the degree of', size: 24, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 2160 },
      children: [new TextRun({ text: degree, size: 30, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 180 },
      children: [new TextRun({ text: 'Submitted by:', size: 24, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 180 },
      children: [new TextRun({ text: student, size: 26, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 1080 },
      children: [new TextRun({ text: studentDegree, size: 24, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 180 },
      children: [new TextRun({ text: 'Submitted to:', size: 24, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 180 },
      children: [new TextRun({ text: adviserName, size: 26, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 0 },
      children: [new TextRun({ text: adviserTitle, size: 24, font: 'Times New Roman', bold: false })],
    }),
  ];
}

function buildAcknowledgementPage(data: any): Paragraph[] {
  const student = data.studentName || '';
  const degree = data.degreeProgram || '';
  const university = 'Jose Rizal Memorial State University';
  const paragraphs: Paragraph[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 360, after: 480 },
      children: [new TextRun({ text: 'ACKNOWLEDGEMENT', bold: true, size: 28, font: 'Times New Roman' })],
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
          spacing: { before: 120, after: 120, line: 360 },
          children: [new TextRun({ text: pText, size: 24, font: 'Times New Roman' })],
        })
      );
    });
  } else {
    paragraphs.push(
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        indent: { firstLine: 720 },
        spacing: { before: 120, after: 120, line: 360 },
        children: [new TextRun({ text: 'No acknowledgement provided.', size: 24, font: 'Times New Roman' })],
      })
    );
  }

  // Signature Block matching the uploaded image format exactly
  paragraphs.push(
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { before: 720, after: 60 },
      children: [new TextRun({ text: student, bold: true, size: 24, font: 'Times New Roman' })],
    }),
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { after: 60 },
      children: [new TextRun({ text: degree, size: 24, font: 'Times New Roman' })],
    }),
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { after: 0 },
      children: [new TextRun({ text: university, size: 24, font: 'Times New Roman' })],
    })
  );

  return paragraphs;
}

function buildSectionPage(title: string, sections: SectionData[]): Paragraph[] {
  const children: any[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 180 },
      children: [new TextRun({ text: title, bold: true, size: 28, font: 'Times New Roman' })],
    }),
  ];

  sections.forEach((section) => {
    const text = typeof section.content === 'string' ? section.content : Array.isArray(section.content) ? section.content.join('\n') : String(section.content ?? '');
    const contentLines = ensureArray(text);

    if (section.title) {
      children.push(
        new Paragraph({
          spacing: { before: 120, after: 60 },
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
            spacing: { after: 60 },
            children: [new TextRun({ text: line, size: 24, font: 'Times New Roman' })],
          })
        );
      });
    } else if (contentLines.length > 0) {
      const combinedText = contentLines.join(' ');
      children.push(
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 80 },
          children: [new TextRun({ text: combinedText, size: 24, font: 'Times New Roman' })],
        })
      );
    }
  });

  return children;
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
              children: [new TextRun({ text: detail_1, size: 20, font: 'Times New Roman' })],
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
              children: [new TextRun({ text: detail_2, size: 20, font: 'Times New Roman' })],
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
          children: [new TextRun({ text: detail, size: 22, font: 'Times New Roman' })],
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
    width: { size: 100, type: WidthType.PERCENTAGE },
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
            margins: { top: 50, bottom: 50, left: 25, right: 25 },
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                spacing: { after: 50, line: 276, lineRule: 'auto' },
                children: [new TextRun({ text: preparedByLabel, bold: true, size: 22, font: 'Times New Roman' })],
              })
            ],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 50, bottom: 50, left: 25, right: 25 },
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                spacing: { after: 50, line: 276, lineRule: 'auto' },
                children: [new TextRun({ text: checkedByLabel, bold: true, size: 22, font: 'Times New Roman' })],
              })
            ],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 25, bottom: 50, left: 25, right: 25 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 25, after: 50, line: 276, lineRule: 'auto' },
                children: [new TextRun({ text: preparedByName, bold: true, size: 22, font: 'Times New Roman' })],
              })
            ],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 25, bottom: 50, left: 25, right: 25 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 25, after: 50, line: 276, lineRule: 'auto' },
                children: [new TextRun({ text: checkedByName, bold: true, size: 22, font: 'Times New Roman' })],
              })
            ],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 50, bottom: 100, left: 50, right: 50 },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 25, after: 50, line: 276, lineRule: 'auto' }, children: [new TextRun({ text: '', size: 22, font: 'Times New Roman' })] })],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 25, bottom: 50, left: 25, right: 25 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 25, after: 50, line: 276, lineRule: 'auto' },
                children: [new TextRun({ text: checkedByRole, size: 22, font: 'Times New Roman' })],
              })
            ],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 25, bottom: 50, left: 25, right: 25 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 25, after: 50, line: 276, lineRule: 'auto' },
                children: [
                  new TextRun({ text: `${dateLabel} `, bold: true, size: 22, font: 'Times New Roman' }),
                  new TextRun({ text: dateValue, bold: true, size: 22, font: 'Times New Roman' }),
                ],
              })
            ],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 25, bottom: 50, left: 25, right: 25 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 25, after: 50, line: 276, lineRule: 'auto' },
                children: [new TextRun({ text: officeInCharge, size: 22, font: 'Times New Roman' })],
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
      children: [new TextRun({ text: 'APPENDICES', bold: true, size: 28, font: 'Times New Roman' })],
    }),
  ];

  if (!appendicesData || !appendicesData.dailyJournal || !Array.isArray(appendicesData.dailyJournal)) return children;

  const layout = appendicesData.dailyJournalLayout === 'report' ? 'report' : 'current';

  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 240, after: 120 },
      children: [new TextRun({ text: 'DAILY WORK ACTIVITIES', bold: true, size: 24, font: 'Times New Roman' })],
    })
  );

  appendicesData.dailyJournal.forEach((weekData: DailyJournalWeek) => {
    if (layout === 'report') {
      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
          children: [new TextRun({ text: `WEEK ${weekData.weekNumber}`, bold: true, size: 24, font: 'Times New Roman' })],
        }),
        buildWeeklyReportTable(weekData.activities, weekData.images || [])
      );

      children.push(
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { before: 120, after: 240 },
          children: [new TextRun({ text: weekData.narrative || 'No narrative provided.', size: 22, font: 'Times New Roman' })],
        })
      );

      return;
    }

    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 120 },
        children: [new TextRun({ text: `WEEK ${weekData.weekNumber}`, bold: true, size: 24, font: 'Times New Roman' })],
      }),
      buildCurrentWeeklyTable(weekData.activities, weekData.totalHours),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 120 },
        children: [new TextRun({ text: `Table ${weekData.weekNumber}: Week ${weekData.weekNumber}`, bold: true, size: 20, font: 'Times New Roman' })],
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 240 },
        children: [new TextRun({ text: weekData.narrative || 'No narrative provided.', size: 24, font: 'Times New Roman' })],
      })
    );

    if (weekData.images && weekData.images.length > 0) {
      children.push(buildImageGridTable(weekData.images));
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

    children.push(new Paragraph({ pageBreakBefore: true }));
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 240, after: 240 },
        children: [new TextRun({ text: app.title, bold: true, size: 28, font: 'Times New Roman' })],
      })
    );

    if (app.isText) {
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
  const tableRows = [
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'DAY', bold: true, size: 22 })] })], width: { size: 20, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'DATE', bold: true, size: 22 })] })], width: { size: 25, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'DAILY ACCOMPLISHMENT', bold: true, size: 22 })] })], width: { size: 40, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'NO. OF WORKING HOURS', bold: true, size: 22 })] })], width: { size: 15, type: WidthType.PERCENTAGE } }),
      ],
    }),
  ];

  const days = activities && activities.length > 0 ? activities : [
    { day: 'Day 1', date: '', accomplishment: '', hours: '' }
  ];

  days.forEach((act) => {
    tableRows.push(
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: act.day || '', size: 22 })] })] }),
          new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: act.date || '', size: 22 })] })] }),
          new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: act.accomplishment || '', size: 22 })] })] }),
          new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(act.hours || ''), size: 22 })] })] }),
        ],
      })
    );
  });

  tableRows.push(
    new TableRow({
      children: [
        new TableCell({ 
          columnSpan: 3, 
          children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: 'TOTAL NUMBER OF HOURS', bold: true, size: 22 })] })] 
        }),
        new TableCell({ 
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(totalHours), bold: true, size: 22 })] })] 
        }),
      ],
    })
  );

  return new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE } });
}

function buildWeeklyReportTable(activities: DailyActivity[], images: (string | AppendixImage)[] = []) {
  const tableRows = [
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'DATE', bold: true, size: 20, font: 'Times New Roman' })] })], width: { size: 20, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'ACCOMPLISHMENT', bold: true, size: 20, font: 'Times New Roman' })] })], width: { size: 45, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'DOCUMENTATION', bold: true, size: 20, font: 'Times New Roman' })] })], width: { size: 35, type: WidthType.PERCENTAGE } }),
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
            spacing: { after: 60 },
            children: [new ImageRun({ type: 'png', data: buf, transformation: { width: 200, height: 135 } })],
          })
        );

        if (detail) {
          imageCellChildren.push(
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 20 },
              children: [new TextRun({ text: detail, size: 16, font: 'Times New Roman' })],
            })
          );
        }
      }
    }

    tableRows.push(
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: act.date || act.day || '', size: 18, font: 'Times New Roman' })] })], verticalAlign: 'center' }),
          new TableCell({ 
            children: (() => {
              const accomplishmentText = act.accomplishment || 'No accomplishment added.';
              const lines = accomplishmentText.split('\n').filter(line => line.trim());
              if (lines.length === 0) {
                return [new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 40 }, children: [new TextRun({ text: 'No accomplishment added.', size: 18, font: 'Times New Roman' })] })];
              }
              return lines.map((line, idx) => 
                new Paragraph({ 
                  alignment: AlignmentType.LEFT, 
                  spacing: { after: idx === lines.length - 1 ? 40 : 20 },
                  bullet: { level: 0 },
                  children: [new TextRun({ text: line.trim(), size: 18, font: 'Times New Roman' })] 
                })
              );
            })()
          }),
          new TableCell({ children: imageCellChildren.length > 0 ? imageCellChildren : [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '', size: 18, font: 'Times New Roman' })] })] }),
        ],
      })
    );
  });

  return new Table({
    rows: tableRows,
    width: { size: 100, type: WidthType.PERCENTAGE },
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
