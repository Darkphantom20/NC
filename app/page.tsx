'use client';

import { useState } from 'react';

const initialState = {
  studentName: 'Ian P. Padilla',
  degreeProgram: 'Bachelor of Science in Information System',
  acknowledgement: 'With deepest gratitude and appreciation, I humbly extend my sincere thanks to all who contributed to my OJT. ',
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
      const response = await fetch('/api/generate-report', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(form),
      });

      if (!response.ok) {
        throw new Error('Failed to generate document');
      }

      const blob = await response.blob();
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
