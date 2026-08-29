"use client";

import { useState } from 'react';

const sectionOrder = [
  'Acknowledgement',
  'Introduction',
  'Organization Analysis',
  'Tasks & Duties',
  'Case Analysis',
  'Reflections',
];

const defaultForm = {
  studentName: 'Ian P. Padilla',
  degreeProgram: 'Bachelor of Science in Information System',
  acknowledgement: 'With deepest gratitude and appreciation, I humbly extend my sincere thanks to all who contributed to my OJT experience and helped me grow in both technical and professional knowledge.',
  background: 'The Management Information Systems Office (MISO) serves as the core technological backbone, responsible for managing, maintaining, and securing the digital infrastructure, information systems, and network communications of the institution.',
  vision: 'To be a premier provider of innovative, reliable, and secure technological solutions and digital services.',
  mission: 'To empower the organization through efficient IT infrastructure, responsive technical support, and robust systems development.',
  objectives: 'To streamline administrative processes and ensure system reliability.',
  coreValues: 'Integrity, Excellence, and Commitment to Public Service.',
  services: 'Provide reliable technical support and ICT services to the institution and its stakeholders.',
  strengths: 'The organization possesses highly skilled and dedicated technical personnel capable of handling complex network and software demands.',
  weaknesses: 'Internal operations occasionally face limitations due to aging hardware upgrades and constrained resource allocations.',
  opportunities: 'There are significant external prospects for adopting advanced cloud solutions and process automation technologies.',
  threats: 'External risks such as evolving cybersecurity vulnerabilities and potential data breaches remain constant challenges.',
  recommendations: 'It is highly recommended to upgrade legacy infrastructure and conduct continuous staff training programs.',
  tasks: 'Assigned tasks included troubleshooting office hardware issues, maintaining network cables, and updating internal database records.',
  procedures: 'Followed strict IT ticketing protocols, observed safety compliance guidelines during hardware repairs, and conformed to daily attendance logs.',
  issue1: 'Encountered unexpected network downtime during a critical system update, disrupting office communications.',
  issue1Action: 'Conducted a line check, isolated the faulty switch port, and switched the main office routing temporarily to a secondary backup line.',
  issue2: 'Faced compatibility errors when deploying a legacy database script onto the updated server environment.',
  issue2Action: 'Debugged SQL syntax constraints, updated driver packages, and successfully refactored connection strings.',
  lessons: 'These situations taught the vital importance of systematic troubleshooting, maintaining proper system backups, and remaining calm under pressure.',
  selfEvaluation: 'The OJT journey served as a transformative learning process, pushing me to transition from theoretical classroom knowledge to practical, fast-paced technical execution.',
  relevancy: 'The host organization directly aligns with my degree program, allowing me to fulfill my expected professional goals of mastering enterprise systems administration and IT support workflows.',
};

const fieldClass =
  'mt-2 w-full rounded-2xl border border-slate-700 bg-slate-950/80 px-4 py-3 text-slate-100 placeholder:text-slate-500 focus:border-cyan-400 focus:outline-none focus:ring-2 focus:ring-cyan-500/30';

export default function Home() {
  const [form, setForm] = useState(defaultForm);
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState('Draft ready');

  const updateField = (key: keyof typeof defaultForm, value: string) => {
    setForm((prev) => ({ ...prev, [key]: value }));
  };

  const handleGenerate = async () => {
    setLoading(true);
    setStatus('Generating DOCX...');

    try {
      const response = await fetch('/api/report', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(form),
      });

      if (!response.ok) {
        throw new Error('Failed to generate the report');
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = 'JRMSU_Narrative_Report.docx';
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);

      setStatus('Download started');
    } catch (error) {
      setStatus('Something went wrong');
      console.error(error);
    } finally {
      setLoading(false);
    }
  };

  return (
    <main className="min-h-screen bg-slate-950 px-6 py-10 text-slate-100">
      <div className="mx-auto max-w-7xl">
        <header className="mb-10 overflow-hidden rounded-[28px] border border-cyan-500/20 bg-slate-900/80 p-8 shadow-glow backdrop-blur-xl">
          <div className="flex flex-col gap-6 md:flex-row md:items-center md:justify-between">
            <div>
              <p className="text-sm uppercase tracking-[0.35em] text-cyan-300">JRMSU</p>
              <h1 className="mt-3 text-4xl font-black tracking-tight text-white md:text-6xl">
                Narrative Report Generator
              </h1>
            </div>
            <div className="rounded-2xl border border-cyan-500/30 bg-cyan-500/10 px-4 py-3 text-right">
              <p className="text-xs uppercase tracking-[0.3em] text-cyan-200">Status</p>
              <p className="text-xl font-bold text-cyan-300">{status}</p>
            </div>
          </div>

          <p className="mt-6 max-w-3xl text-lg text-slate-300">
            A realistic internship report builder designed to turn your OJT experiences into a polished narrative report with chapter-based guidance and professional Word export.
          </p>
        </header>

        <section className="mb-10 grid gap-6 md:grid-cols-3">
          <Card label="Sections" value="6" color="cyan" />
          <Card label="Format" value="DOCX" color="violet" />
          <Card label="Mode" value="Web" color="emerald" />
        </section>

        <div className="grid gap-8 lg:grid-cols-[1.4fr_0.6fr]">
          <form className="space-y-8 rounded-[28px] border border-slate-800 bg-slate-900/80 p-6">
            <SectionBlock title="1. Acknowledgement">
              <label className="block text-sm font-medium text-slate-200">
                Student Name
                <input value={form.studentName} onChange={(e) => updateField('studentName', e.target.value)} className={fieldClass} />
              </label>
              <label className="mt-4 block text-sm font-medium text-slate-200">
                Degree Program
                <input value={form.degreeProgram} onChange={(e) => updateField('degreeProgram', e.target.value)} className={fieldClass} />
              </label>
              <label className="mt-4 block text-sm font-medium text-slate-200">
                Acknowledgement Paragraph
                <textarea value={form.acknowledgement} onChange={(e) => updateField('acknowledgement', e.target.value)} rows={4} className={fieldClass} />
              </label>
            </SectionBlock>

            <SectionBlock title="2. Introduction">
              <label className="block text-sm font-medium text-slate-200">
                Background of Organization
                <textarea value={form.background} onChange={(e) => updateField('background', e.target.value)} rows={4} className={fieldClass} />
              </label>
              <div className="mt-4 grid gap-4 md:grid-cols-2">
                <label className="block text-sm font-medium text-slate-200">
                  Vision
                  <input value={form.vision} onChange={(e) => updateField('vision', e.target.value)} className={fieldClass} />
                </label>
                <label className="block text-sm font-medium text-slate-200">
                  Mission
                  <input value={form.mission} onChange={(e) => updateField('mission', e.target.value)} className={fieldClass} />
                </label>
              </div>
              <div className="mt-4 grid gap-4 md:grid-cols-3">
                <label className="block text-sm font-medium text-slate-200">
                  Objectives
                  <textarea value={form.objectives} onChange={(e) => updateField('objectives', e.target.value)} rows={3} className={fieldClass} />
                </label>
                <label className="block text-sm font-medium text-slate-200">
                  Core Values
                  <textarea value={form.coreValues} onChange={(e) => updateField('coreValues', e.target.value)} rows={3} className={fieldClass} />
                </label>
                <label className="block text-sm font-medium text-slate-200">
                  Services Offered
                  <textarea value={form.services} onChange={(e) => updateField('services', e.target.value)} rows={3} className={fieldClass} />
                </label>
              </div>
            </SectionBlock>

            <SectionBlock title="3. Organization Analysis">
              <div className="grid gap-4 md:grid-cols-2">
                <label className="block text-sm font-medium text-slate-200">
                  Strengths
                  <textarea value={form.strengths} onChange={(e) => updateField('strengths', e.target.value)} rows={4} className={fieldClass} />
                </label>
                <label className="block text-sm font-medium text-slate-200">
                  Weaknesses
                  <textarea value={form.weaknesses} onChange={(e) => updateField('weaknesses', e.target.value)} rows={4} className={fieldClass} />
                </label>
                <label className="block text-sm font-medium text-slate-200">
                  Opportunities
                  <textarea value={form.opportunities} onChange={(e) => updateField('opportunities', e.target.value)} rows={4} className={fieldClass} />
                </label>
                <label className="block text-sm font-medium text-slate-200">
                  Threats
                  <textarea value={form.threats} onChange={(e) => updateField('threats', e.target.value)} rows={4} className={fieldClass} />
                </label>
              </div>
              <label className="mt-4 block text-sm font-medium text-slate-200">
                Recommendations
                <textarea value={form.recommendations} onChange={(e) => updateField('recommendations', e.target.value)} rows={4} className={fieldClass} />
              </label>
            </SectionBlock>

            <SectionBlock title="4. Tasks and Duties">
              <label className="block text-sm font-medium text-slate-200">
                Assigned Tasks
                <textarea value={form.tasks} onChange={(e) => updateField('tasks', e.target.value)} rows={4} className={fieldClass} />
              </label>
              <label className="mt-4 block text-sm font-medium text-slate-200">
                Procedures Conformed
                <textarea value={form.procedures} onChange={(e) => updateField('procedures', e.target.value)} rows={4} className={fieldClass} />
              </label>
            </SectionBlock>

            <SectionBlock title="5. Case Analysis">
              <div className="grid gap-4 md:grid-cols-2">
                <label className="block text-sm font-medium text-slate-200">
                  Problem 1
                  <textarea value={form.issue1} onChange={(e) => updateField('issue1', e.target.value)} rows={4} className={fieldClass} />
                </label>
                <label className="block text-sm font-medium text-slate-200">
                  Solution/Action 1
                  <textarea value={form.issue1Action} onChange={(e) => updateField('issue1Action', e.target.value)} rows={4} className={fieldClass} />
                </label>
                <label className="block text-sm font-medium text-slate-200">
                  Problem 2
                  <textarea value={form.issue2} onChange={(e) => updateField('issue2', e.target.value)} rows={4} className={fieldClass} />
                </label>
                <label className="block text-sm font-medium text-slate-200">
                  Solution/Action 2
                  <textarea value={form.issue2Action} onChange={(e) => updateField('issue2Action', e.target.value)} rows={4} className={fieldClass} />
                </label>
              </div>
              <label className="mt-4 block text-sm font-medium text-slate-200">
                Lessons Learned
                <textarea value={form.lessons} onChange={(e) => updateField('lessons', e.target.value)} rows={4} className={fieldClass} />
              </label>
            </SectionBlock>

            <SectionBlock title="6. Reflections">
              <label className="block text-sm font-medium text-slate-200">
                Self-Evaluation
                <textarea value={form.selfEvaluation} onChange={(e) => updateField('selfEvaluation', e.target.value)} rows={4} className={fieldClass} />
              </label>
              <label className="mt-4 block text-sm font-medium text-slate-200">
                Relevancy to Course and Goals
                <textarea value={form.relevancy} onChange={(e) => updateField('relevancy', e.target.value)} rows={4} className={fieldClass} />
              </label>
            </SectionBlock>
          </form>

          <aside className="flex flex-col gap-4 rounded-[28px] border border-slate-800 bg-gradient-to-br from-cyan-500/10 via-slate-900 to-violet-500/10 p-6">
            <h3 className="text-xl font-bold text-white">Report Preview</h3>

            <div className="space-y-4 text-sm text-slate-300">
              <Box label="Student" value={form.studentName} />
              <Box label="Program" value={form.degreeProgram} />
              <Box label="Organization" value={form.background.slice(0, 90) + (form.background.length > 90 ? '...' : '')} />
              <Box label="Status" value={status} />
            </div>

            <div className="mt-auto rounded-2xl border border-slate-700 bg-slate-950/70 p-4">
              <p className="text-xs uppercase tracking-[0.2em] text-slate-400">Current flow</p>
              <div className="mt-4 space-y-3">
                {sectionOrder.map((section, index) => (
                  <div key={section} className="flex items-center gap-3">
                    <span className="flex h-7 w-7 items-center justify-center rounded-full bg-cyan-500/15 text-xs font-bold text-cyan-300">
                      {index + 1}
                    </span>
                    <span className="text-slate-200">{section}</span>
                  </div>
                ))}
              </div>
            </div>

            <button
              type="button"
              onClick={handleGenerate}
              disabled={loading}
              className="mt-4 rounded-2xl bg-gradient-to-r from-cyan-500 to-violet-500 px-4 py-3 text-base font-semibold text-white shadow-glow transition hover:opacity-95 disabled:cursor-not-allowed disabled:opacity-60"
            >
              {loading ? 'Generating...' : 'Generate Report'}
            </button>
          </aside>
        </div>
      </div>
    </main>
  );
}

function SectionBlock({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <section className="rounded-3xl border border-slate-800 bg-slate-950/60 p-5">
      <h2 className="mb-4 text-xl font-bold text-white">{title}</h2>
      {children}
    </section>
  );
}

function Card({ label, value, color }: { label: string; value: string; color: 'cyan' | 'violet' | 'emerald' }) {
  const tone = {
    cyan: 'text-cyan-300',
    violet: 'text-violet-300',
    emerald: 'text-emerald-300',
  }[color];

  return (
    <div className="rounded-2xl border border-slate-700 bg-slate-900/80 p-5">
      <p className="text-sm uppercase tracking-[0.2em] text-slate-400">{label}</p>
      <p className={`mt-4 text-4xl font-bold ${tone}`}>{value}</p>
    </div>
  );
}

function Box({ label, value }: { label: string; value: string }) {
  return (
    <div className="rounded-2xl border border-slate-700 bg-slate-950/70 p-4">
      <p className="text-xs uppercase tracking-[0.2em] text-slate-400">{label}</p>
      <p className="mt-2 text-base font-semibold text-white">{value}</p>
    </div>
  );
}
