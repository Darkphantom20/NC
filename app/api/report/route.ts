import fs from 'fs';
import path from 'path';
import { createClient } from '@supabase/supabase-js';
import { Document, Packer, Paragraph, TextRun, AlignmentType, ImageRun, LineRuleType } from 'docx';

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
    // If it's a newline-separated string or comma-separated, handle gracefully or return as single item array
    return val.split('\n').map(v => v.trim()).filter(Boolean);
  }
  return [];
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
                top: 576,    // 0.4 inches
                bottom: 1440, // 1.0 inch
                left: 1152,  // 0.8 inches
                right: 1152  // 0.8 inches
              },
            },
          },
          children: [
            // Section 1: Acknowledgement
            ...buildAcknowledgementPage(data.studentName, data.degreeProgram, data.acknowledgement),
            
            // Section 2: Introduction
            new Paragraph({ pageBreakBefore: true }),
            ...buildSectionPage('2. INTRODUCTION', [
              ['Background of the Organization', ensureArray(data.background).length ? ensureArray(data.background) : [data.background || 'The Management Information Systems Office (MISO) serves as the core technological backbone...']],
              ['Vision', ensureArray(data.vision).length ? ensureArray(data.vision) : [data.vision || 'To be a premier provider of innovative, reliable, and secure technological solutions...']],
              ['Mission', ensureArray(data.mission).length ? ensureArray(data.mission) : [data.mission || 'To empower the organization through efficient IT infrastructure...']],
              ['Objectives', ensureArray(data.objectives).length ? ensureArray(data.objectives) : ['To streamline administrative processes and ensure system reliability.']],
              ['Core Values', ensureArray(data.coreValues).length ? ensureArray(data.coreValues) : ['Integrity, Excellence, and Commitment to Public Service.']],
              ['Products and Services Offered', ensureArray(data.services).length ? ensureArray(data.services) : ['Provide reliable technical support and ICT services.']],
            ], true),

            // Section 3: Company Analysis
            new Paragraph({ pageBreakBefore: true }),
            ...buildSectionPage('3. ORGANIZATION / COMPANY ANALYSIS', [
              ['Strengths of the Organization (Internal factors)', ensureArray(data.strengths).length ? ensureArray(data.strengths) : ['The organization possesses highly skilled and dedicated technical personnel...']],
              ['Weaknesses of the Organization (Internal factors)', ensureArray(data.weaknesses).length ? ensureArray(data.weaknesses) : ['Internal operations occasionally face limitations due to aging hardware...']],
              ['Opportunities of the Organization (External factors)', ensureArray(data.opportunities).length ? ensureArray(data.opportunities) : ['There are significant external prospects for adopting advanced cloud solutions...']],
              ['Threats of the Organization (External factors)', ensureArray(data.threats).length ? ensureArray(data.threats) : ['External risks such as evolving cybersecurity vulnerabilities remain constant challenges.']],
              ['Recommendations for Improvement', ensureArray(data.recommendations).length ? ensureArray(data.recommendations) : ['It is highly recommended to upgrade legacy infrastructure and conduct continuous staff training.']],
            ]),

            // Section 4: Tasks and Duties
            new Paragraph({ pageBreakBefore: true }),
            ...buildSectionPage('4. TASKS AND DUTIES', [
              ['Assigned Tasks and Responsibilities', ensureArray(data.tasks).length ? ensureArray(data.tasks) : ['Assigned tasks included troubleshooting office hardware issues and maintaining network cables.']],
              ['Duties and Procedures Conformed', ensureArray(data.procedures).length ? ensureArray(data.procedures) : ['Followed strict IT ticketing protocols and observed safety compliance guidelines.']],
            ]),

            // Section 5: Case Analysis
            new Paragraph({ pageBreakBefore: true }),
            ...buildSectionPage('5. CASE ANALYSIS', [
              ['Issue / Problem 1', [
                `Description: ${data.issue1 || 'Encountered unexpected network downtime during a critical system update.'}`,
                `Strategy/Action Undertaken: ${data.issue1Action || 'Conducted a line check, isolated the faulty switch port, and switched routing to backup.'}`
              ]],
              ['Issue / Problem 2', [
                `Description: ${data.issue2 || 'Faced compatibility errors when deploying a legacy database script.'}`,
                `Strategy/Action Undertaken: ${data.issue2Action || 'Debugged SQL syntax constraints and updated driver packages.'}`
              ]],
              ['Lessons Learned from the Situations', ensureArray(data.lessons).length ? ensureArray(data.lessons) : ['These situations taught the vital importance of systematic troubleshooting and backups.']],
            ]),

            // Section 6: Reflections
            new Paragraph({ pageBreakBefore: true }),
            ...buildSectionPage('6. REFLECTIONS', [
              ['Self-Evaluation from the Learning Process Experienced', ensureArray(data.selfEvaluation).length ? ensureArray(data.selfEvaluation) : ['The OJT journey served as a transformative learning process bridging theory and execution.']],
              ['Relevancy of the Organization with Your Programme of Study and Expected Goals', ensureArray(data.relevancy).length ? ensureArray(data.relevancy) : ['The host organization directly aligns with my degree program goals of mastering enterprise administration.']],
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
  const name = `${studentName || 'IAN P. PADILLA'}`.toUpperCase();
  const program = degreeProgram || 'Bachelor of Science in Information System';
  const ackParas = ensureArray(acknowledgement).length ? ensureArray(acknowledgement) : [
    typeof acknowledgement === 'string' && acknowledgement.trim() ? acknowledgement : 'With deepest gratitude and appreciation, I humbly extend my sincere thanks to all who contributed to my OJT.'
  ];

  const children: any[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 20, after: 120 },
      children: buildHeaderImage(),
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 259, after: 172 }, // Pt(18) before, Pt(12) after approx
      children: [
        new TextRun({ text: 'ACKNOWLEDGEMENT', bold: true, size: 28, font: 'Times New Roman' }),
      ],
    }),
  ];

  for (const text of ackParas) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.JUSTIFY,
        spacing: { line: 276, lineRule: LineRuleType.AUTO, after: 86 }, // 1.15 line spacing, 6pt after
        children: [
          new TextRun({ text, size: 22, font: 'Times New Roman' }),
        ],
      })
    );
  }

  children.push(
    new Paragraph({ spacing: { after: 144 } }), // empty space
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { line: 276, lineRule: LineRuleType.AUTO },
      children: [
        new TextRun({ text: `${name}\n`, bold: true, size: 22, font: 'Times New Roman' }),
        new TextRun({ text: `${program}\n`, size: 22, font: 'Times New Roman' }),
        new TextRun({ text: 'Jose Rizal Memorial State University', size: 22, font: 'Times New Roman' }),
      ],
    })
  );

  return children;
}

function buildSectionPage(sectionTitle: string, sections: Array<[string, string[]]>, isBulletListCategory = false) {
  const children: any[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 20, after: 120 },
      children: buildHeaderImage(),
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 259, after: 201 }, // Pt(18) before, Pt(14) after approx
      children: [
        new TextRun({ text: sectionTitle, bold: true, size: 28, font: 'Times New Roman' }),
      ],
    }),
  ];

  for (const [subtitle, lines] of sections) {
    children.push(
      new Paragraph({
        spacing: { before: 144, after: 57 }, // Pt(10) before, Pt(4) after
        children: [
          new TextRun({ text: subtitle, bold: true, size: 24, font: 'Times New Roman' }),
        ],
      })
    );

    const useBullets = isBulletListCategory && ['Objectives', 'Core Values', 'Products and Services Offered'].includes(subtitle) && lines.length > 1;

    for (const line of lines) {
      children.push(
        new Paragraph({
          alignment: AlignmentType.JUSTIFY,
          bullet: useBullets ? { level: 0 } : undefined,
          spacing: { line: 276, lineRule: LineRuleType.AUTO, after: 57 },
          children: [
            new TextRun({ text: line, size: 22, font: 'Times New Roman' }),
          ],
        })
      );
    }
  }

  return children;
}