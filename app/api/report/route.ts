import fs from 'fs';
import path from 'path';
import { createClient } from '@supabase/supabase-js';
import { Document, Packer, Paragraph, TextRun, AlignmentType, ImageRun } from 'docx';

export const runtime = 'nodejs';

const OUTPUT_FILE_NAME = process.env.OUTPUT_FILE_NAME || 'JRMSU_Exact_Template.docx';
const HEADER_IMAGE_PATH = path.join(process.cwd(), 'public', 'Picture1.png');
const HEADER_IMAGE = fs.existsSync(HEADER_IMAGE_PATH) ? fs.readFileSync(HEADER_IMAGE_PATH) : null;
const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;
const hasSupabaseConfig = Boolean(supabaseUrl && supabaseKey);

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
        degree_program: data.degreeProgram || null,
        acknowledgement: data.acknowledgement || null,
        organization_background: data.background || null,
        vision: data.vision || null,
        mission: data.mission || null,
        objectives: data.objectives || null,
        core_values: data.coreValues || null,
        services: data.services || null,
        strengths: data.strengths || null,
        weaknesses: data.weaknesses || null,
        opportunities: data.opportunities || null,
        threats: data.threats || null,
        recommendations: data.recommendations || null,
        tasks: data.tasks || null,
        procedures: data.procedures || null,
        issue_1_description: data.issue1 || null,
        issue_1_strategy: data.issue1Action || null,
        issue_2_description: data.issue2 || null,
        issue_2_strategy: data.issue2Action || null,
        lessons: data.lessons || null,
        self_evaluation: data.selfEvaluation || null,
        relevancy: data.relevancy || null,
      };

      const { error } = await supabase.from('reports').insert([payload]);
      if (error) {
        console.error('Supabase insert error:', error);
      }
    } else {
      console.warn('Supabase environment variables are not configured. Report data was not stored.');
    }

    const doc = new Document({
      sections: [
        {
          properties: {
            page: {
              margin: { top: 720, right: 720, bottom: 720, left: 720 },
            },
          },
          children: [
            ...buildAcknowledgementPage(data.studentName, data.degreeProgram, data.acknowledgement),
            new Paragraph({ pageBreakBefore: true }),
            ...buildSectionPage('2. INTRODUCTION', [
              ['Background of the Organization', data.background],
              ['Vision', data.vision],
              ['Mission', data.mission],
              ['Objectives', data.objectives],
              ['Core Values', data.coreValues],
              ['Products and Services Offered', data.services],
            ]),
            ...buildSectionPage('3. ORGANIZATION / COMPANY ANALYSIS', [
              ['Strengths of the Organization (Internal factors)', data.strengths],
              ['Weaknesses of the Organization (Internal factors)', data.weaknesses],
              ['Opportunities of the Organization (External factors)', data.opportunities],
              ['Threats of the Organization (External factors)', data.threats],
              ['Recommendations for Improvement', data.recommendations],
            ]),
            ...buildSectionPage('4. TASKS AND DUTIES', [
              ['Assigned Tasks and Responsibilities', data.tasks],
              ['Duties and Procedures Conformed', data.procedures],
            ]),
            ...buildSectionPage('5. CASE ANALYSIS', [
              ['Issue / Problem 1', `${data.issue1 || 'No issue provided.'}`],
              ['Strategy/Action Undertaken for Problem 1', `${data.issue1Action || 'No action provided.'}`],
              ['Issue / Problem 2', `${data.issue2 || 'No issue provided.'}`],
              ['Strategy/Action Undertaken for Problem 2', `${data.issue2Action || 'No action provided.'}`],
              ['Lessons Learned from the Situations', data.lessons],
            ]),
            ...buildSectionPage('6. REFLECTIONS', [
              ['Self-Evaluation from the Learning Process Experienced', data.selfEvaluation],
              ['Relevancy of the Organization with Your Programme of Study and Expected Goals', data.relevancy],
            ]),
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
      headers: { 'Content-Type': 'application/json' },
    });
  }
}

function buildAcknowledgementPage(studentName: string, degreeProgram: string, acknowledgement: string) {
  const name = `${studentName || 'Student Name'}`.toUpperCase();
  const program = degreeProgram || 'Degree Program';
  const text = acknowledgement || 'No acknowledgement provided.';

  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 20 },
      children: [
        ...(HEADER_IMAGE
          ? [
              new ImageRun({
                data: HEADER_IMAGE,
                transformation: { width: 140, height: 70 },
              }),
            ]
          : []),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 120 },
      children: [
        new TextRun({ text: 'ACKNOWLEDGEMENT', bold: true, size: 28, font: 'Times New Roman' }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { after: 100 },
      children: [
        new TextRun({ text, size: 22, font: 'Times New Roman' }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 80, after: 10 },
      children: [
        new TextRun({ text: name, bold: true, size: 24, font: 'Times New Roman' }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0 },
      children: [
        new TextRun({ text: program, size: 18, font: 'Times New Roman' }),
      ],
    }),
  ];
}

function buildSectionPage(sectionTitle: string, sections: Array<[string, string]>) {
  const children: any[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 14 },
      children: [
        ...(HEADER_IMAGE
          ? [
              new ImageRun({
                data: HEADER_IMAGE,
                transformation: { width: 140, height: 70 },
              }),
            ]
          : []),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 80, after: 120 },
      children: [
        new TextRun({ text: sectionTitle, bold: true, size: 28, font: 'Times New Roman' }),
      ],
    }),
  ];

  for (const [title, content] of sections) {
    children.push(
      new Paragraph({
        spacing: { before: 100, after: 40 },
        children: [
          new TextRun({ text: title, bold: true, size: 22, font: 'Times New Roman' }),
        ],
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 80 },
        children: [
          new TextRun({ text: content || 'No information provided.', size: 20, font: 'Times New Roman' }),
        ],
      }),
    );
  }

  return children;
}
