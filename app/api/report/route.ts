import fs from 'fs';
import path from 'path';
import { createClient } from '@supabase/supabase-js';
import { Document, Packer, Paragraph, TextRun, AlignmentType, ImageRun, Header } from 'docx';

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
      transformation: { width: 500, height: 100 },
    }),
  ];
}

const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;
const hasSupabaseConfig = Boolean(supabaseUrl && supabaseKey);

function ensureArray(val: any): string[] {
  if (Array.isArray(val)) return val.map(v => String(v).trim()).filter(Boolean);
  if (typeof val === 'string' && val.trim().length > 0) {
    return val.split('\n').map(v => v.trim()).filter(Boolean);
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
        degree_program: data.degreeProgram || null,
        acknowledgement: JSON.stringify(data.acknowledgement) || null,
        organization_background: data.background || null,
        vision: data.vision || null,
        mission: data.mission || null,
        objectives: JSON.stringify(data.objectives) || null,
        core_values: JSON.stringify(data.coreValues) || null,
        services: JSON.stringify(data.services) || null,
        strengths: JSON.stringify(data.strengths) || null,
        weaknesses: JSON.stringify(data.weaknesses) || null,
        opportunities: JSON.stringify(data.opportunities) || null,
        threats: JSON.stringify(data.threats) || null,
        recommendations: JSON.stringify(data.recommendations) || null,
        tasks: JSON.stringify(data.tasks) || null,
        procedures: JSON.stringify(data.procedures) || null,
        issue_1_description: data.issue1 || null,
        issue_1_strategy: data.issue1Action || null,
        issue_2_description: data.issue2 || null,
        issue_2_strategy: data.issue2Action || null,
        lessons: JSON.stringify(data.lessons) || null,
        self_evaluation: JSON.stringify(data.selfEvaluation) || null,
        relevancy: JSON.stringify(data.relevancy) || null,
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
              margin: { 
                top: 1440,    // 1.0 inch for standard balanced margins
                bottom: 1440, 
                left: 1440,  
                right: 1440,
                header: 720,  // 0.5 inch margin for the header
              },
            },
          },
          // Embed the header at the document section level
          headers: {
            default: new Header({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: buildHeaderImage(),
                }),
              ],
            }),
          },
          children: [
            // Section 1: Acknowledgement
            ...buildAcknowledgementPage(data.studentName, data.degreeProgram, data.acknowledgement),
            
            // Section 2: Introduction
            new Paragraph({ pageBreakBefore: true }),
            ...buildSectionPage('2. INTRODUCTION', [
              { title: 'Background of the Organization', content: data.background || 'The Management Information Systems Office (MISO) serves as the core technological backbone...' },
              { title: 'Vision', content: data.vision || 'A technologically advanced and efficient Provincial Government...' },
              { title: 'Mission', content: data.mission || 'To provide reliable information systems...' },
              { title: 'Objectives', content: data.objectives, isBullet: true },
              { title: 'Core Values', content: data.coreValues, isBullet: true },
              { title: 'Products and Services Offered', content: data.services || 'Provide reliable technical support and ICT services.' },
            ]),

            // Section 3: Company Analysis
            new Paragraph({ pageBreakBefore: true }),
            ...buildSectionPage('3. ORGANIZATION / COMPANY ANALYSIS', [
              { title: 'Strengths of the Organization (Internal factors)', content: data.strengths, isBullet: true },
              { title: 'Weaknesses of the Organization (Internal factors)', content: data.weaknesses, isBullet: true },
              { title: 'Opportunities of the Organization (External factors)', content: data.opportunities, isBullet: true },
              { title: 'Threats of the Organization (External factors)', content: data.threats, isBullet: true },
              { title: 'Recommendations for Improvement', content: data.recommendations, isBullet: true },
            ]),

            // Section 4: Tasks and Duties
            new Paragraph({ pageBreakBefore: true }),
            ...buildSectionPage('4. TASKS AND DUTIES', [
              { title: 'Assigned Tasks and Responsibilities', content: data.tasks, isBullet: true },
              { title: 'Duties and Procedures Conformed', content: data.procedures, isBullet: true },
            ]),

            // Section 5: Case Analysis
            new Paragraph({ pageBreakBefore: true }),
            ...buildSectionPage('5. CASE ANALYSIS', [
              { title: 'Issue / Problem 1', content: data.issue1 || 'No issue provided.' },
              { title: 'Strategy/Action Undertaken for Problem 1', content: data.issue1Action || 'No action provided.' },
              { title: 'Issue / Problem 2', content: data.issue2 || 'No issue provided.' },
              { title: 'Strategy/Action Undertaken for Problem 2', content: data.issue2Action || 'No action provided.' },
              { title: 'Lessons Learned from the Situations', content: data.lessons },
            ]),

            // Section 6: Reflections
            new Paragraph({ pageBreakBefore: true }),
            ...buildSectionPage('6. REFLECTIONS', [
              { title: 'Self-Evaluation from the Learning Process Experienced', content: data.selfEvaluation },
              { title: 'Relevancy of the Organization with Your Programme of Study and Expected Goals', content: data.relevancy },
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

function buildAcknowledgementPage(studentName: string, degreeProgram: string, acknowledgement: any) {
  const name = `${studentName || 'HAYNA G. DAUD'}`.toUpperCase();
  const program = degreeProgram || 'Bachelor of Science in Information System';
  
  const ackParas = ensureArray(acknowledgement).length ? ensureArray(acknowledgement) : [
    'With deepest gratitude and appreciation, I humbly extend my sincere thanks to all who contributed to my OJT experience and helped me grow in both technical and professional knowledge.'
  ];

  const children: any[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 360 }, 
      children: [
        new TextRun({ text: 'ACKNOWLEDGEMENT', bold: true, size: 28, font: 'Times New Roman' }),
      ],
    }),
  ];

  for (const text of ackParas) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.JUSTIFY,
        spacing: { after: 200 }, 
        children: [
          new TextRun({ text, size: 24, font: 'Times New Roman' }), 
        ],
      })
    );
  }

  children.push(
    new Paragraph({ spacing: { before: 480 } }), 
    new Paragraph({
      alignment: AlignmentType.LEFT,
      children: [
        new TextRun({ text: name, bold: true, size: 24, font: 'Times New Roman' }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.LEFT,
      children: [
        new TextRun({ text: program, size: 24, font: 'Times New Roman' }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.LEFT,
      children: [
        new TextRun({ text: 'Jose Rizal Memorial State University', size: 24, font: 'Times New Roman' }),
      ],
    })
  );

  return children;
}

function buildSectionPage(sectionTitle: string, sections: SectionData[]) {
  const children: any[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 240 },
      children: [
        new TextRun({ text: sectionTitle, bold: true, size: 28, font: 'Times New Roman' }),
      ],
    }),
  ];

  for (const section of sections) {
    children.push(
      new Paragraph({
        spacing: { before: 240, after: 120 }, 
        children: [
          new TextRun({ text: section.title, bold: true, size: 24, font: 'Times New Roman' }),
        ],
      })
    );

    const lines = ensureArray(section.content);
    if (lines.length === 0) {
      lines.push('No information provided.');
    }

    for (const line of lines) {
      const cleanLine = line.replace(/^[\*\-\•]\s*/, ''); 
      
      children.push(
        new Paragraph({
          alignment: section.isBullet ? AlignmentType.LEFT : AlignmentType.JUSTIFY,
          bullet: section.isBullet ? { level: 0 } : undefined,
          spacing: { after: 120 },
          children: [
            new TextRun({ text: cleanLine, size: 24, font: 'Times New Roman' }), 
          ],
        })
      );
    }
  }

  return children;
}