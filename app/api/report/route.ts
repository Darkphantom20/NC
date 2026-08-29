import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } from 'docx';

export async function POST(request: Request) {
  try {
    const data = await request.json();

    if (!data || typeof data !== 'object') {
      return new Response(JSON.stringify({ error: 'Invalid payload' }), {
        status: 400,
        headers: { 'Content-Type': 'application/json' },
      });
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
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 240 },
              children: [
                new TextRun({ text: 'ACKNOWLEDGEMENT', bold: true, size: 30 }),
              ],
            }),
            new Paragraph({
              spacing: { after: 200 },
              children: [new TextRun({ text: data.acknowledgement || 'No acknowledgement provided.', size: 22 })],
            }),
            new Paragraph({ children: [new TextRun({ text: `${data.studentName || 'Student Name'}`.toUpperCase(), bold: true, size: 22 })] }),
            new Paragraph({ children: [new TextRun({ text: data.degreeProgram || 'Degree Program', size: 22 })] }),

            new Paragraph({ pageBreakBefore: true, alignment: AlignmentType.CENTER, spacing: { before: 240, after: 200 }, children: [new TextRun({ text: '2. INTRODUCTION', bold: true, size: 30 })] }),
            ...makeSection('Background of the Organization', data.background),
            ...makeSection('Vision', data.vision),
            ...makeSection('Mission', data.mission),
            ...makeSection('Objectives', data.objectives),
            ...makeSection('Core Values', data.coreValues),
            ...makeSection('Products and Services Offered', data.services),

            new Paragraph({ pageBreakBefore: true, alignment: AlignmentType.CENTER, spacing: { before: 240, after: 200 }, children: [new TextRun({ text: '3. ORGANIZATION / COMPANY ANALYSIS', bold: true, size: 30 })] }),
            ...makeSection('Strengths of the Organization (Internal factors)', data.strengths),
            ...makeSection('Weaknesses of the Organization (Internal factors)', data.weaknesses),
            ...makeSection('Opportunities of the Organization (External factors)', data.opportunities),
            ...makeSection('Threats of the Organization (External factors)', data.threats),
            ...makeSection('Recommendations for Improvement', data.recommendations),

            new Paragraph({ pageBreakBefore: true, alignment: AlignmentType.CENTER, spacing: { before: 240, after: 200 }, children: [new TextRun({ text: '4. TASKS AND DUTIES', bold: true, size: 30 })] }),
            ...makeSection('Assigned Tasks and Responsibilities', data.tasks),
            ...makeSection('Duties and Procedures Conformed', data.procedures),

            new Paragraph({ pageBreakBefore: true, alignment: AlignmentType.CENTER, spacing: { before: 240, after: 200 }, children: [new TextRun({ text: '5. CASE ANALYSIS', bold: true, size: 30 })] }),
            ...makeSection('Issue / Problem 1', data.issue1),
            ...makeSection('Strategy/Action Undertaken for Problem 1', data.issue1Action),
            ...makeSection('Issue / Problem 2', data.issue2),
            ...makeSection('Strategy/Action Undertaken for Problem 2', data.issue2Action),
            ...makeSection('Lessons Learned from the Situations', data.lessons),

            new Paragraph({ pageBreakBefore: true, alignment: AlignmentType.CENTER, spacing: { before: 240, after: 200 }, children: [new TextRun({ text: '6. REFLECTIONS', bold: true, size: 30 })] }),
            ...makeSection('Self-Evaluation from the Learning Process Experienced', data.selfEvaluation),
            ...makeSection('Relevancy of the Organization with Your Programme of Study and Expected Goals', data.relevancy),
          ],
        },
      ],
    });

    const buffer = await Packer.toBuffer(doc);

    return new Response(buffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': 'attachment; filename="JRMSU_Narrative_Report.docx"',
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

function makeSection(title: string, content: string) {
  return [
    new Paragraph({
      heading: HeadingLevel.HEADING_3,
      spacing: { before: 180, after: 80 },
      children: [new TextRun({ text: title, bold: true, size: 24 })],
    }),
    new Paragraph({
      spacing: { after: 80 },
      children: [new TextRun({ text: content || 'No information provided.', size: 22 })],
    }),
  ];
}
