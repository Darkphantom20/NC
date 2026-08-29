const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, PageBreak } = require('docx');

function normalizeParagraphs(value) {
  if (Array.isArray(value)) {
    return value.filter(item => item && item.toString().trim().length > 0);
  }
  return value && value.toString().trim().length > 0 ? [value.toString()] : [];
}

function createBodyParagraph(text, options = {}) {
  const para = new Paragraph({
    alignment: options.alignment ?? AlignmentType.JUSTIFY,
    spacing: {
      after: 120,
      before: options.before ?? 0,
      line: 360,
    },
    indent: options.indent,
    bullet: options.bullet ?? undefined,
  });

  if (typeof text === 'string') {
    para.addChildElement(
      new TextRun({
        text,
        font: 'Times New Roman',
        size: 22,
        bold: options.bold ?? false,
      }),
    );
  }

  return para;
}

function buildReportDocument(data) {
  const doc = new Document({
    styles: {
      paragraphStyles: [
        {
          id: 'Title',
          name: 'Title',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            size: 28,
            bold: true,
            font: 'Times New Roman',
          },
        },
        {
          id: 'Subtitle',
          name: 'Subtitle',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            size: 24,
            bold: true,
            font: 'Times New Roman',
          },
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 400,
              bottom: 1000,
              left: 800,
              right: 800,
            },
          },
          header: {
            distance: 200,
          },
        },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 180, after: 120 },
            children: [
              new TextRun({ text: 'ACKNOWLEDGEMENT', bold: true, size: 28, font: 'Times New Roman' }),
            ],
          }),
          ...normalizeParagraphs(data.acknowledgement).map((item) => createBodyParagraph(item)),
          new Paragraph({
            spacing: { before: 120 },
            children: [
              new TextRun({ text: `${String(data.studentName || 'Student Name').toUpperCase()}\n`, bold: true, font: 'Times New Roman', size: 22 }),
              new TextRun({ text: `${data.degreeProgram || 'Degree Program'}\n`, font: 'Times New Roman', size: 22 }),
              new TextRun({ text: 'Jose Rizal Memorial State University', font: 'Times New Roman', size: 22 }),
            ],
          }),
          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 180, after: 120 },
            children: [
              new TextRun({ text: '2. INTRODUCTION', bold: true, size: 28, font: 'Times New Roman' }),
            ],
          }),
          createBodyParagraph('Background of the Organization', { bold: true, before: 100 }),
          createBodyParagraph(data.organizationBackground || 'No background provided.'),
          createBodyParagraph('Vision', { bold: true, before: 100 }),
          createBodyParagraph(data.vision || 'No vision provided.'),
          createBodyParagraph('Mission', { bold: true, before: 100 }),
          createBodyParagraph(data.mission || 'No mission provided.'),
          createBodyParagraph('Objectives', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.objectives).map((item) => createBodyParagraph(item, { bullet: true, indent: { left: 720 } })),
          createBodyParagraph('Core Values', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.coreValues).map((item) => createBodyParagraph(item, { bullet: true, indent: { left: 720 } })),
          createBodyParagraph('Products and Services Offered', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.services).map((item) => createBodyParagraph(item, { bullet: true, indent: { left: 720 } })),
          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 180, after: 120 },
            children: [
              new TextRun({ text: '3. ORGANIZATION / COMPANY ANALYSIS', bold: true, size: 28, font: 'Times New Roman' }),
            ],
          }),
          createBodyParagraph('Strengths of the Organization (Internal factors)', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.strengths).map((item) => createBodyParagraph(item)),
          createBodyParagraph('Weaknesses of the Organization (Internal factors)', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.weaknesses).map((item) => createBodyParagraph(item)),
          createBodyParagraph('Opportunities of the Organization (External factors)', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.opportunities).map((item) => createBodyParagraph(item)),
          createBodyParagraph('Threats of the Organization (External factors)', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.threats).map((item) => createBodyParagraph(item)),
          createBodyParagraph('Recommendations for Improvement', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.recommendations).map((item) => createBodyParagraph(item)),
          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 180, after: 120 },
            children: [
              new TextRun({ text: '4. TASKS AND DUTIES', bold: true, size: 28, font: 'Times New Roman' }),
            ],
          }),
          createBodyParagraph('Assigned Tasks and Responsibilities', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.tasks).map((item) => createBodyParagraph(item)),
          createBodyParagraph('Duties and Procedures Conformed', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.procedures).map((item) => createBodyParagraph(item)),
          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 180, after: 120 },
            children: [
              new TextRun({ text: '5. CASE ANALYSIS', bold: true, size: 28, font: 'Times New Roman' }),
            ],
          }),
          createBodyParagraph('Issue / Problem 1', { bold: true, before: 100 }),
          createBodyParagraph(`Description: ${data.issue1Description || 'No issue description provided.'}`),
          createBodyParagraph(`Strategy/Action Undertaken: ${data.issue1Strategy || 'No strategy provided.'}`),
          createBodyParagraph('Issue / Problem 2', { bold: true, before: 100 }),
          createBodyParagraph(`Description: ${data.issue2Description || 'No issue description provided.'}`),
          createBodyParagraph(`Strategy/Action Undertaken: ${data.issue2Strategy || 'No strategy provided.'}`),
          createBodyParagraph('Lessons Learned from the Situations', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.lessons).map((item) => createBodyParagraph(item)),
          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 180, after: 120 },
            children: [
              new TextRun({ text: '6. REFLECTIONS', bold: true, size: 28, font: 'Times New Roman' }),
            ],
          }),
          createBodyParagraph('Self-Evaluation from the Learning Process Experienced', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.selfEvaluation).map((item) => createBodyParagraph(item)),
          createBodyParagraph('Relevancy of the Organization with Your Programme of Study and Expected Goals', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.relevancy).map((item) => createBodyParagraph(item)),
        ],
      },
    ],
  });

  return doc;
}

module.exports = {
  buildReportDocument,
};
