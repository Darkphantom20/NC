'use client';

import { useState } from 'react';
import { AlignmentType, Document, PageBreak, Paragraph, Packer, TextRun } from 'docx';

const initialState = {
  studentName: 'Ian P. Padilla',
  degreeProgram: 'Bachelor of Science in Information System',
  acknowledgement: 'With deepest gratitude and appreciation, I humbly extend my sincere thanks to all who contributed to my OJT.',
  organizationBackground: 'The Management Information Systems Office (MISO) serves as the core technological backbone, responsible for managing, maintaining, and securing the digital infrastructure, information systems, and network communications of the institution.',
  vision: 'To be a premier provider of innovative, reliable, and secure technological solutions and digital services.',
  mission: 'To empower the organization through efficient IT infrastructure, responsive technical support, and robust systems development.',
  objectives: 'To streamline administrative processes and ensure system reliability.',
  coreValues: 'Integrity, Excellence, and Commitment to Public Service.',
  services: 'Provide reliable technical support and ICT services.',
  strengths: 'The organization possesses highly skilled and dedicated technical personnel capable of handling complex network and software demands.',
  weaknesses: 'Internal operations occasionally face limitations due to aging hardware upgrades and constrained resource allocations.',
  opportunities: 'There are significant external prospects for adopting advanced cloud solutions and process automation technologies.',
  threats: 'External risks such as evolving cybersecurity vulnerabilities and potential data breaches remain constant challenges.',
  recommendations: 'It is highly recommended to upgrade legacy infrastructure and conduct continuous staff training programs.',
  tasks: 'Assigned tasks included troubleshooting office hardware issues, maintaining network cables, and updating internal database records.',
  procedures: 'Followed strict IT ticketing protocols, observed safety compliance guidelines during hardware repairs, and conformed to daily attendance logs.',
  issue1Description: 'Encountered unexpected network downtime during a critical system update, disrupting office communications.',
  issue1Strategy: 'Conducted a line check, isolated the faulty switch port, and switched the main office routing temporarily to a secondary backup line.',
  issue2Description: 'Faced compatibility errors when deploying a legacy database script onto the updated server environment.',
  issue2Strategy: 'Debugged SQL syntax constraints, updated driver packages, and successfully refactored connection strings.',
  lessons: 'These situations taught the vital importance of systematic troubleshooting, maintaining proper system backups, and remaining calm under pressure.',
  selfEvaluation: 'The OJT journey served as a transformative learning process, pushing me to transition from theoretical classroom knowledge to practical, fast-paced technical execution.',
  relevancy: 'The host organization directly aligns with my degree program, allowing me to fulfill my expected professional goals of mastering enterprise systems administration and IT support workflows.',
};

const fields = [
  ['studentName', 'Student Name'],
  ['degreeProgram', 'Degree Program'],
  ['acknowledgement', 'Acknowledgement'],
  ['organizationBackground', 'Background of the Organization'],
  ['vision', 'Vision'],
  ['mission', 'Mission'],
  ['objectives', 'Objectives'],
  ['coreValues', 'Core Values'],
  ['services', 'Products and Services Offered'],
  ['strengths', 'Strengths'],
  ['weaknesses', 'Weaknesses'],
  ['opportunities', 'Opportunities'],
  ['threats', 'Threats'],
  ['recommendations', 'Recommendations'],
  ['tasks', 'Assigned Tasks'],
  ['procedures', 'Duties and Procedures'],
  ['issue1Description', 'Issue 1 Description'],
  ['issue1Strategy', 'Issue 1 Strategy'],
  ['issue2Description', 'Issue 2 Description'],
  ['issue2Strategy', 'Issue 2 Strategy'],
  ['lessons', 'Lessons Learned'],
  ['selfEvaluation', 'Self Evaluation'],
  ['relevancy', 'Relevancy'],
] as const;

const normalizeParagraphs = (value: string | string[] | undefined) => {
  if (Array.isArray(value)) {
    return value.filter((item) => item && item.trim().length > 0);
  }

  return value && value.trim().length > 0 ? [value] : [];
};

const createParagraph = (text: string, options: { bold?: boolean; before?: number; bullet?: boolean; indent?: number } = {}) => {
  const para = new Paragraph({
    alignment: AlignmentType.JUSTIFY,
    spacing: { after: 120, before: options.before ?? 0, line: 360 },
    indent: options.indent ? { left: options.indent } : undefined,
    bullet: options.bullet ? { level: 0 } : undefined,
  });

  para.addChildElement(
    new TextRun({
      text,
      font: 'Times New Roman',
      size: 22,
      bold: options.bold ?? false,
    }),
  );

  return para;
};

const buildReportDocument = (data: typeof initialState) => {
  const doc = new Document({
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
        },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 180, after: 120 },
            children: [new TextRun({ text: 'ACKNOWLEDGEMENT', bold: true, size: 28, font: 'Times New Roman' })],
          }),
          ...normalizeParagraphs(data.acknowledgement).map((item) => createParagraph(item)),
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
            children: [new TextRun({ text: '2. INTRODUCTION', bold: true, size: 28, font: 'Times New Roman' })],
          }),
          createParagraph('Background of the Organization', { bold: true, before: 100 }),
          createParagraph(data.organizationBackground || 'No background provided.'),
          createParagraph('Vision', { bold: true, before: 100 }),
          createParagraph(data.vision || 'No vision provided.'),
          createParagraph('Mission', { bold: true, before: 100 }),
          createParagraph(data.mission || 'No mission provided.'),
          createParagraph('Objectives', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.objectives).map((item) => createParagraph(item, { bullet: true, indent: 720 })),
          createParagraph('Core Values', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.coreValues).map((item) => createParagraph(item, { bullet: true, indent: 720 })),
          createParagraph('Products and Services Offered', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.services).map((item) => createParagraph(item, { bullet: true, indent: 720 })),
          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 180, after: 120 },
            children: [new TextRun({ text: '3. ORGANIZATION / COMPANY ANALYSIS', bold: true, size: 28, font: 'Times New Roman' })],
          }),
          createParagraph('Strengths of the Organization (Internal factors)', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.strengths).map((item) => createParagraph(item)),
          createParagraph('Weaknesses of the Organization (Internal factors)', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.weaknesses).map((item) => createParagraph(item)),
          createParagraph('Opportunities of the Organization (External factors)', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.opportunities).map((item) => createParagraph(item)),
          createParagraph('Threats of the Organization (External factors)', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.threats).map((item) => createParagraph(item)),
          createParagraph('Recommendations for Improvement', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.recommendations).map((item) => createParagraph(item)),
          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 180, after: 120 },
            children: [new TextRun({ text: '4. TASKS AND DUTIES', bold: true, size: 28, font: 'Times New Roman' })],
          }),
          createParagraph('Assigned Tasks and Responsibilities', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.tasks).map((item) => createParagraph(item)),
          createParagraph('Duties and Procedures Conformed', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.procedures).map((item) => createParagraph(item)),
          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 180, after: 120 },
            children: [new TextRun({ text: '5. CASE ANALYSIS', bold: true, size: 28, font: 'Times New Roman' })],
          }),
          createParagraph('Issue / Problem 1', { bold: true, before: 100 }),
          createParagraph(`Description: ${data.issue1Description || 'No issue description provided.'}`),
          createParagraph(`Strategy/Action Undertaken: ${data.issue1Strategy || 'No strategy provided.'}`),
          createParagraph('Issue / Problem 2', { bold: true, before: 100 }),
          createParagraph(`Description: ${data.issue2Description || 'No issue description provided.'}`),
          createParagraph(`Strategy/Action Undertaken: ${data.issue2Strategy || 'No strategy provided.'}`),
          createParagraph('Lessons Learned from the Situations', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.lessons).map((item) => createParagraph(item)),
          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 180, after: 120 },
            children: [new TextRun({ text: '6. REFLECTIONS', bold: true, size: 28, font: 'Times New Roman' })],
          }),
          createParagraph('Self-Evaluation from the Learning Process Experienced', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.selfEvaluation).map((item) => createParagraph(item)),
          createParagraph('Relevancy of the Organization with Your Programme of Study and Expected Goals', { bold: true, before: 100 }),
          ...normalizeParagraphs(data.relevancy).map((item) => createParagraph(item)),
        ],
      },
    ],
  });

  return doc;
};

export default function Page() {
  const [form, setForm] = useState(initialState);
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState('');

  const handleChange = (key: keyof typeof initialState, value: string) => {
    setForm((prev) => ({ ...prev, [key]: value }));
  };

  const handleSubmit = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setLoading(true);
    setStatus('Generating report...');

    try {
      const doc = buildReportDocument(form);
      const blob = await Packer.toBlob(doc);
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'JRMSU_Narrative_Report_Sections_1_to_6.docx';
      a.click();
      window.URL.revokeObjectURL(url);
      setStatus('Document downloaded successfully.');
    } catch (error) {
      setStatus('Something went wrong while generating the document.');
      console.error(error);
    } finally {
      setLoading(false);
    }
  };

  return (
    <main style={{ maxWidth: 1100, margin: '0 auto', padding: '2rem 1rem 4rem' }}>
      <h1 style={{ textAlign: 'center', marginBottom: '1.5rem' }}>JRMSU Narrative Report Generator</h1>

      <form onSubmit={handleSubmit}>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(300px, 1fr))', gap: '1rem' }}>
          {fields.map(([key, label]) => (
            <label key={key} style={{ display: 'flex', flexDirection: 'column', gap: '0.4rem', fontWeight: 600 }}>
              {label}
              <textarea
                value={String(form[key])}
                onChange={(e) => handleChange(key, e.target.value)}
                rows={4}
                style={{ padding: '0.75rem', borderRadius: '8px', border: '1px solid #ccc', resize: 'vertical' }}
              />
            </label>
          ))}
        </div>

        <div style={{ marginTop: '1.5rem', textAlign: 'center' }}>
          <button type="submit" disabled={loading} style={{ padding: '0.8rem 1.5rem', borderRadius: '8px', border: 'none', background: '#0f172a', color: '#fff', cursor: 'pointer' }}>
            {loading ? 'Generating...' : 'Generate Word Document'}
          </button>
        </div>

        {status ? <p style={{ textAlign: 'center', marginTop: '1rem' }}>{status}</p> : null}
      </form>
    </main>
  );
}
