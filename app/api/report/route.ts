import fs from 'fs';
import path from 'path';
import { createClient } from '@supabase/supabase-js';
import { Document, Packer, Paragraph, TextRun, AlignmentType, ImageRun, Header, Table, TableRow, TableCell, WidthType, PageBreak } from 'docx';

export const runtime = 'nodejs';

const OUTPUT_FILE_NAME = process.env.OUTPUT_FILE_NAME || 'JRMSU_Exact_Template.docx';
const HEADER_IMAGE_PATH = path.join(process.cwd(), 'public', 'Picture1.png');
const HEADER_IMAGE = fs.existsSync(HEADER_IMAGE_PATH) ? fs.readFileSync(HEADER_IMAGE_PATH) : null;

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

interface SectionData {
  title: string;
  content: any;
  isBullet?: boolean;
}

export async function POST(request: Request) {
  try {
    const data = await request.json();

    if (!data || typeof data !== 'object') {
      return new Response(JSON.stringify({ error: 'Invalid payload' }), {
        status: 400,
        headers: { 'Content-Type': 'application/json' },
      });
    }

    if (hasSupabaseConfig) {
      const supabase = createClient(supabaseUrl!, supabaseKey!);
      const payload = {
        student_name: data.studentName || null,
        // ... (Keep existing Supabase mapping logic)
        appendices: JSON.stringify(data.appendices) || null,
      };

      const { error } = await supabase.from('reports').insert([payload]);
      if (error) console.error('Supabase insert error:', error);
    }

    const doc = new Document({
      sections: [
        {
          properties: {
            page: {
              size: { width: 12240, height: 20160 },
              margin: { top: 1440, bottom: 1440, left: 1440, right: 1440, header: 0 },
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
            // Sections 1-6 (Keep existing implementation)
            ...buildAcknowledgementPage(data.studentName, data.degreeProgram, data.acknowledgement),
            new Paragraph({ pageBreakBefore: true }),
            ...buildSectionPage('2. INTRODUCTION', [
              { title: 'Background of the Organization', content: data.background },
              { title: 'Vision', content: data.vision },
              { title: 'Mission', content: data.mission },
              { title: 'Objectives', content: data.objectives, isBullet: true },
              { title: 'Core Values', content: data.coreValues, isBullet: true },
              { title: 'Products and Services Offered', content: data.services },
            ]),
            new Paragraph({ pageBreakBefore: true }),
            ...buildSectionPage('3. ORGANIZATION / COMPANY ANALYSIS', [
              { title: 'Strengths', content: data.strengths, isBullet: true },
              { title: 'Weaknesses', content: data.weaknesses, isBullet: true },
              { title: 'Opportunities', content: data.opportunities, isBullet: true },
              { title: 'Threats', content: data.threats, isBullet: true },
              { title: 'Recommendations for Improvement', content: data.recommendations, isBullet: true },
            ]),
            new Paragraph({ pageBreakBefore: true }),
            ...buildSectionPage('4. TASKS AND DUTIES', [
              { title: 'Assigned Tasks and Responsibilities', content: data.tasks, isBullet: true },
              { title: 'Duties and Procedures Conformed', content: data.procedures, isBullet: true },
            ]),
            new Paragraph({ pageBreakBefore: true }),
            ...buildSectionPage('5. CASE ANALYSIS', [
              { title: 'Issue / Problem 1', content: data.issue1 },
              { title: 'Strategy/Action Undertaken for Problem 1', content: data.issue1Action },
              { title: 'Issue / Problem 2', content: data.issue2 },
              { title: 'Strategy/Action Undertaken for Problem 2', content: data.issue2Action },
              { title: 'Lessons Learned from the Situations', content: data.lessons },
            ]),
            new Paragraph({ pageBreakBefore: true }),
            ...buildSectionPage('6. REFLECTIONS', [
              { title: 'Self-Evaluation', content: data.selfEvaluation },
              { title: 'Relevancy of the Organization', content: data.relevancy },
            ]),
            
            // Section 7: Appendices
            new Paragraph({ pageBreakBefore: true }),
            ...buildAppendicesPage(data.appendices),
          ],
        },
      ],
    });

    const buffer = await Packer.toBuffer(doc);

    return new Response(buffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="${OUTPUT_FILE_NAME}"`,
      },
    });
  } catch (error) {
    console.error(error);
    return new Response(JSON.stringify({ error: 'Something went wrong while generating the report.' }), {
      status: 500,
    });
  }
}

function buildAcknowledgementPage(studentName: string, degreeProgram: string, acknowledgement: string): Paragraph[] {
  const children: any[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 200 },
      children: [new TextRun({ text: 'ACKNOWLEDGEMENT', bold: true, size: 28, font: 'Times New Roman' })],
    }),
    new Paragraph({
      spacing: { before: 80, after: 80 },
      children: [new TextRun({ text: `Student Name: ${studentName || 'N/A'}`, size: 24, font: 'Times New Roman' })],
    }),
    new Paragraph({
      spacing: { before: 80, after: 80 },
      children: [new TextRun({ text: `Degree Program: ${degreeProgram || 'N/A'}`, size: 24, font: 'Times New Roman' })],
    }),
    new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      indent: { left: 720 },
      spacing: { before: 120, after: 120 },
      children: [new TextRun({ text: acknowledgement || 'No acknowledgement provided.', size: 24, font: 'Times New Roman' })],
    }),
  ];

  return children;
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
    const text = typeof section.content === 'string'
      ? section.content
      : Array.isArray(section.content)
        ? section.content.join('\n')
        : String(section.content ?? '');

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

function buildAppendicesPage(appendicesData: any) {
  const children: any[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 360 },
      children: [
        new TextRun({ text: '7. APPENDICES', bold: true, size: 28, font: 'Times New Roman' }),
      ],
    }),
  ];

  if (!appendicesData) return children;

  // 1. Daily Journal
  if (appendicesData.dailyJournal && Array.isArray(appendicesData.dailyJournal)) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 240, after: 120 },
        children: [new TextRun({ text: 'DAILY WORK ACTIVITIES', bold: true, size: 24, font: 'Times New Roman' })],
      })
    );

    appendicesData.dailyJournal.forEach((weekData: any) => {
      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
          children: [new TextRun({ text: `WEEK ${weekData.weekNumber}`, bold: true, size: 24, font: 'Times New Roman' })],
        }),
        buildWeeklyTable(weekData.activities, weekData.totalHours),
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

      // Attach 1-3 Images for the Daily Report
      if (weekData.images && Array.isArray(weekData.images)) {
        weekData.images.slice(0, 3).forEach((imgBase64: string) => {
          children.push(
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 120 },
              children: [
                new ImageRun({
                  type: 'png',
                  data: Buffer.from(imgBase64.replace(/^data:image\/\w+;base64,/, ""), 'base64'),
                  transformation: { width: 400, height: 300 },
                }),
              ],
            })
          );
        });
      }
    });
  }

  // Helper to add standard appendix documents (Certificates, Resume, etc.)
  const appendixList = [
    { title: 'Certificate of Participation (PRIME Seminar)', key: 'certParticipation' },
    { title: 'One page narrative report of PRIME', key: 'primeNarrative', isText: true },
    { title: 'Resume and Application Letter', key: 'resume' },
    { title: 'Evaluation of Grades & Validated Enrolment Form', key: 'grades' },
    { title: 'Medical Certificate & Parent\'s Waiver', key: 'medicalWaiver' },
    { title: 'Letter of Acceptance', key: 'letterAcceptance' },
    { title: 'Daily Time Record (DTR) - hours rendered', key: 'dtr' },
    { title: 'Student’s Intern Rating Sheet & Performance Evaluation', key: 'ratingSheet' },
    { title: 'Certificate of Completion & Memorandum of Agreement', key: 'certCompletion' },
  ];

  appendixList.forEach(app => {
    children.push(new Paragraph({ pageBreakBefore: true }));
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 240, after: 240 },
        children: [new TextRun({ text: app.title, bold: true, size: 28, font: 'Times New Roman' })],
      })
    );

    const data = appendicesData[app.key];
    if (data) {
      if (app.isText) {
        children.push(
          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            children: [new TextRun({ text: data, size: 24, font: 'Times New Roman' })],
          })
        );
      } else {
        children.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new ImageRun({
                type: 'png',
                data: Buffer.from(data.replace(/^data:image\/\w+;base64,/, ""), 'base64'),
                transformation: { width: 550, height: 750 },
              }),
            ],
          })
        );
      }
    } else {
      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: '[Document not provided]', italics: true, size: 24, font: 'Times New Roman' })],
        })
      );
    }
  });

  return children;
}

function buildWeeklyTable(activities: any[], totalHours: number | string) {
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

  const defaultDays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
  
  defaultDays.forEach((day, index) => {
    const act = activities && activities[index] ? activities[index] : { date: '', accomplishment: '', hours: '' };
    tableRows.push(
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: day, size: 22 })] })] }),
          new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: act.date, size: 22 })] })] }),
          new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: act.accomplishment, size: 22 })] })] }),
          new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(act.hours), size: 22 })] })] }),
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

  return new Table({
    rows: tableRows,
    width: { size: 100, type: WidthType.PERCENTAGE },
  });
}