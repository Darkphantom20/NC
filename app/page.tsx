"use client";

import { useState } from 'react';

const sectionOrder = [
  'Acknowledgement',
  'Introduction',
  'Organization Analysis',
  'Tasks & Duties',
  'Case Analysis',
  'Reflections',
  'Appendices'
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
  
  appendices: {
    dailyJournal: [
      {
        weekNumber: 1,
        totalHours: 40,
        narrative: 'This week focused on orientation and basic network troubleshooting.',
        activities: [
          { day: 'Day 1', date: 'Oct 2', accomplishment: 'Orientation and facility tour', hours: 8 },
          { day: 'Day 2', date: 'Oct 3', accomplishment: 'Network cable termination and testing', hours: 8 },
          { day: 'Day 3', date: 'Oct 4', accomplishment: 'PC hardware cleaning and maintenance', hours: 8 },
          { day: 'Day 4', date: 'Oct 5', accomplishment: 'Operating system updates deployment', hours: 8 },
          { day: 'Day 5', date: 'Oct 6', accomplishment: 'Helpdesk ticketing queue review', hours: 8 }
        ],
        images: [] as { base64: string; detail: string }[]
      }
    ],
    certParticipation: '',
    primeNarrative: 'The PRIME seminar provided profound insights into professional workplace ethics and modern enterprise standards.',
    resume: '',
    grades: '',
    medicalWaiver: '',
    letterAcceptance: '',
    dtr: '',
    ratingSheet: '',
    certCompletion: ''
  }
};

const fieldClass =
  'mt-2 w-full rounded-xl border border-slate-700 bg-slate-950/80 px-3 py-2.5 text-sm text-slate-100 placeholder:text-slate-500 focus:border-cyan-400 focus:outline-none focus:ring-2 focus:ring-cyan-500/30 sm:rounded-2xl sm:px-4 sm:py-3 sm:text-base';

const fileInputClass = 
  'mt-2 block w-full text-sm text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-cyan-500/10 file:text-cyan-400 hover:file:bg-cyan-500/20';

export default function Home() {
  const [form, setForm] = useState(defaultForm);
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState('Draft ready');

  const updateField = (key: keyof typeof defaultForm, value: string) => {
    setForm((prev) => ({ ...prev, [key]: value }));
  };

  const updateAppendixText = (key: keyof typeof defaultForm.appendices, value: string) => {
    setForm((prev) => ({
      ...prev,
      appendices: { ...prev.appendices, [key]: value }
    }));
  };

  // Add / Remove Weeks dynamically
  const addWeek = () => {
    setForm((prev) => {
      const nextWeekNum = prev.appendices.dailyJournal.length + 1;
      const startDayNum = prev.appendices.dailyJournal.reduce((acc, w) => acc + w.activities.length, 0) + 1;
      return {
        ...prev,
        appendices: {
          ...prev.appendices,
          dailyJournal: [
            ...prev.appendices.dailyJournal,
            {
              weekNumber: nextWeekNum,
              totalHours: 40,
              narrative: '',
              activities: Array.from({ length: 5 }, (_, i) => ({
                day: `Day ${startDayNum + i}`,
                date: '',
                accomplishment: '',
                hours: 8
              })),
              images: []
            }
          ]
        }
      };
    });
  };

  const removeWeek = (weekIndex: number) => {
    setForm((prev) => {
      const updated = prev.appendices.dailyJournal.filter((_, idx) => idx !== weekIndex);
      return { ...prev, appendices: { ...prev.appendices, dailyJournal: updated } };
    });
  };

  // Add / Remove Days
  const addDay = (weekIndex: number) => {
    setForm((prev) => {
      const journal = [...prev.appendices.dailyJournal];
      const week = journal[weekIndex];
      const totalDaysOverall = journal.reduce((acc, w) => acc + w.activities.length, 0);
      week.activities.push({
        day: `Day ${totalDaysOverall + 1}`,
        date: '',
        accomplishment: '',
        hours: 8
      });
      week.totalHours = week.activities.reduce((acc, curr) => acc + (Number(curr.hours) || 0), 0);
      return { ...prev, appendices: { ...prev.appendices, dailyJournal: journal } };
    });
  };

  const updateDayActivity = (weekIndex: number, dayIndex: number, key: string, value: any) => {
    setForm((prev) => {
      const journal = [...prev.appendices.dailyJournal];
      const activities = [...journal[weekIndex].activities];
      activities[dayIndex] = { ...activities[dayIndex], [key]: value };
      journal[weekIndex].activities = activities;
      journal[weekIndex].totalHours = activities.reduce((acc, curr) => acc + (Number(curr.hours) || 0), 0);
      return { ...prev, appendices: { ...prev.appendices, dailyJournal: journal } };
    });
  };

  const removeDay = (weekIndex: number, dayIndex: number) => {
    setForm((prev) => {
      const journal = [...prev.appendices.dailyJournal];
      journal[weekIndex].activities = journal[weekIndex].activities.filter((_, idx) => idx !== dayIndex);
      journal[weekIndex].totalHours = journal[weekIndex].activities.reduce((acc, curr) => acc + (Number(curr.hours) || 0), 0);
      return { ...prev, appendices: { ...prev.appendices, dailyJournal: journal } };
    });
  };

  // Handle standard appendix documents (Strictly Images Only)
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>, appendixKey: keyof typeof defaultForm.appendices) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (!file.type.startsWith('image/')) {
      alert('Please upload an image file (PNG, JPG). PDFs are not supported for image rendering.');
      return;
    }

    const reader = new FileReader();
    reader.onload = (event) => {
      const base64 = event.target?.result as string;
      setForm((prev) => ({
        ...prev,
        appendices: { ...prev.appendices, [appendixKey]: base64 }
      }));
    };
    reader.readAsDataURL(file);
  };

  // Base64 image uploader for Daily Journal
  const handleJournalImageUpload = (e: React.ChangeEvent<HTMLInputElement>, weekIndex: number) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (!file.type.startsWith('image/')) {
      alert('Please upload an image file (PNG, JPG).');
      return;
    }

    const reader = new FileReader();
    reader.onload = (event) => {
      const base64 = event.target?.result as string;
      setForm((prev) => {
        const newJournal = [...prev.appendices.dailyJournal];
        if (newJournal[weekIndex].images.length < 3) {
          newJournal[weekIndex].images.push({ base64, detail: `Photo documentation for Week ${newJournal[weekIndex].weekNumber}` });
        } else {
          alert('Maximum of 3 images allowed per week.');
        }
        return { ...prev, appendices: { ...prev.appendices, dailyJournal: newJournal } };
      });
    };
    reader.readAsDataURL(file);
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
        const errData = await response.json();
        throw new Error(errData.error || 'Failed to generate the report');
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
    } catch (error: any) {
      setStatus('Error generating file');
      alert(error.message);
      console.error(error);
    } finally {
      setLoading(false);
    }
  };

  return (
    <main className="min-h-screen bg-slate-950 px-4 py-6 text-slate-100 sm:px-6 sm:py-10 lg:px-8">
      <div className="mx-auto max-w-7xl">
        <header className="mb-6 overflow-hidden rounded-[24px] border border-cyan-500/20 bg-slate-900/80 p-5 shadow-glow backdrop-blur-xl sm:mb-10 sm:rounded-[28px] sm:p-8">
          <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
            <div>
              <p className="text-xs uppercase tracking-[0.25em] text-cyan-300 sm:text-sm">JRMSU</p>
              <h1 className="mt-3 text-3xl font-black tracking-tight text-white sm:text-4xl md:text-6xl">
                Narrative Report Generator
              </h1>
            </div>
            <div className="rounded-2xl border border-cyan-500/30 bg-cyan-500/10 px-3 py-2.5 text-left sm:px-4 sm:py-3 md:text-right">
              <p className="text-[10px] uppercase tracking-[0.25em] text-cyan-200 sm:text-xs">Status</p>
              <p className="mt-1 text-lg font-bold text-cyan-300 sm:text-xl">{status}</p>
            </div>
          </div>
        </header>

        <section className="mb-6 grid gap-4 sm:mb-10 sm:grid-cols-3 sm:gap-6">
          <Card label="Sections" value="7" color="cyan" />
          <Card label="Format" value="DOCX" color="violet" />
          <Card label="Mode" value="Web" color="emerald" />
        </section>

        <div className="grid gap-6 lg:grid-cols-[1.4fr_0.6fr] lg:gap-8">
          <form className="space-y-6 rounded-[24px] border border-slate-800 bg-slate-900/80 p-4 sm:space-y-8 sm:rounded-[28px] sm:p-6">
            
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

            <SectionBlock title="7. Appendices">
              <div className="space-y-8">
                
                {/* Dynamic Daily Journal Section */}
                <div>
                  <div className="flex items-center justify-between mb-4">
                    <h3 className="text-lg font-bold text-cyan-300">Daily Journal (Day 1 to Last Day)</h3>
                    <button
                      type="button"
                      onClick={addWeek}
                      className="rounded-xl bg-cyan-500/20 px-3 py-1.5 text-xs font-semibold text-cyan-300 hover:bg-cyan-500/30"
                    >
                      + Add Week
                    </button>
                  </div>

                  {form.appendices.dailyJournal.map((week, weekIdx) => (
                    <div key={weekIdx} className="mb-6 rounded-2xl border border-slate-800 bg-slate-900/60 p-4 space-y-4">
                      <div className="flex items-center justify-between">
                        <span className="font-bold text-white">WEEK {week.weekNumber} (Total Hours: {week.totalHours})</span>
                        {form.appendices.dailyJournal.length > 1 && (
                          <button
                            type="button"
                            onClick={() => removeWeek(weekIdx)}
                            className="text-xs text-rose-400 hover:underline"
                          >
                            Remove Week
                          </button>
                        )}
                      </div>

                      {/* Day 1 ... Day N Entries */}
                      <div className="space-y-3">
                        <p className="text-xs font-semibold text-slate-400 uppercase tracking-wider">Daily Activities</p>
                        {week.activities.map((act, dayIdx) => (
                          <div key={dayIdx} className="grid gap-2 sm:grid-cols-[1fr_1fr_2.5fr_1fr_auto] items-center bg-slate-950/60 p-2.5 rounded-xl border border-slate-800">
                            <input
                              type="text"
                              placeholder="Day (e.g. Day 1)"
                              value={act.day}
                              onChange={(e) => updateDayActivity(weekIdx, dayIdx, 'day', e.target.value)}
                              className="rounded-lg border border-slate-700 bg-slate-900 px-2 py-1 text-xs text-slate-100"
                            />
                            <input
                              type="text"
                              placeholder="Date"
                              value={act.date}
                              onChange={(e) => updateDayActivity(weekIdx, dayIdx, 'date', e.target.value)}
                              className="rounded-lg border border-slate-700 bg-slate-900 px-2 py-1 text-xs text-slate-100"
                            />
                            <input
                              type="text"
                              placeholder="Daily Accomplishment"
                              value={act.accomplishment}
                              onChange={(e) => updateDayActivity(weekIdx, dayIdx, 'accomplishment', e.target.value)}
                              className="rounded-lg border border-slate-700 bg-slate-900 px-2 py-1 text-xs text-slate-100"
                            />
                            <input
                              type="number"
                              placeholder="Hrs"
                              value={act.hours}
                              onChange={(e) => updateDayActivity(weekIdx, dayIdx, 'hours', e.target.value)}
                              className="rounded-lg border border-slate-700 bg-slate-900 px-2 py-1 text-xs text-slate-100"
                            />
                            <button
                              type="button"
                              onClick={() => removeDay(weekIdx, dayIdx)}
                              className="text-rose-400 text-xs px-2 hover:text-rose-300"
                            >
                              ✕
                            </button>
                          </div>
                        ))}

                        <button
                          type="button"
                          onClick={() => addDay(weekIdx)}
                          className="mt-2 text-xs font-semibold text-cyan-400 hover:underline"
                        >
                          + Add Day to Week {week.weekNumber}
                        </button>
                      </div>

                      {/* Weekly Narrative */}
                      <label className="block text-sm font-medium text-slate-200">
                        Weekly Narrative Report
                        <textarea
                          value={week.narrative}
                          onChange={(e) => {
                            const newJournal = [...form.appendices.dailyJournal];
                            newJournal[weekIdx].narrative = e.target.value;
                            setForm((prev) => ({ ...prev, appendices: { ...prev.appendices, dailyJournal: newJournal } }));
                          }}
                          rows={2}
                          className={fieldClass}
                        />
                      </label>

                      {/* 1-3 Pictures Uploader */}
                      <div>
                        <label className="block text-sm font-medium text-slate-200">
                          Upload Pictures for Week {week.weekNumber} (Max 3)
                          <input type="file" accept="image/png, image/jpeg, image/jpg, image/webp" onChange={(e) => handleJournalImageUpload(e, weekIdx)} className={fileInputClass} />
                        </label>
                        
                        {week.images && week.images.length > 0 && (
                          <div className="mt-3 space-y-2">
                            {week.images.map((img, imgIdx) => (
                              <div key={imgIdx} className="flex items-center gap-3 bg-slate-950 p-2 rounded-xl border border-slate-800">
                                <img src={img.base64} alt="Preview" className="h-10 w-10 object-cover rounded-lg" />
                                <input
                                  type="text"
                                  placeholder="Caption / Picture Detail"
                                  value={img.detail}
                                  onChange={(e) => {
                                    const newJournal = [...form.appendices.dailyJournal];
                                    newJournal[weekIdx].images[imgIdx].detail = e.target.value;
                                    setForm((prev) => ({ ...prev, appendices: { ...prev.appendices, dailyJournal: newJournal } }));
                                  }}
                                  className="w-full rounded-lg border border-slate-700 bg-slate-900 px-2 py-1 text-xs text-slate-100"
                                />
                                <button
                                  type="button"
                                  onClick={() => {
                                    const newJournal = [...form.appendices.dailyJournal];
                                    newJournal[weekIdx].images = newJournal[weekIdx].images.filter((_, idx) => idx !== imgIdx);
                                    setForm((prev) => ({ ...prev, appendices: { ...prev.appendices, dailyJournal: newJournal } }));
                                  }}
                                  className="text-xs text-rose-400 px-2"
                                >
                                  ✕
                                </button>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    </div>
                  ))}
                </div>

                {/* PRIME Narrative */}
                <label className="block text-sm font-medium text-slate-200">
                  One Page Narrative Report of PRIME
                  <textarea value={form.appendices.primeNarrative} onChange={(e) => updateAppendixText('primeNarrative', e.target.value)} rows={3} className={fieldClass} />
                </label>

                {/* Standard Document Uploads */}
                <div className="grid gap-4 md:grid-cols-2">
                  <FileUpload label="Certificate of Participation (PRIME)" onChange={(e) => handleFileUpload(e, 'certParticipation')} />
                  <FileUpload label="Resume and Application Letter" onChange={(e) => handleFileUpload(e, 'resume')} />
                  <FileUpload label="Evaluation of Grades & Validated Enrolment" onChange={(e) => handleFileUpload(e, 'grades')} />
                  <FileUpload label="Medical Certificate & Parent's Waiver" onChange={(e) => handleFileUpload(e, 'medicalWaiver')} />
                  <FileUpload label="Letter of Acceptance" onChange={(e) => handleFileUpload(e, 'letterAcceptance')} />
                  <FileUpload label="Daily Time Record (DTR)" onChange={(e) => handleFileUpload(e, 'dtr')} />
                  <FileUpload label="Intern Rating Sheet & Eval" onChange={(e) => handleFileUpload(e, 'ratingSheet')} />
                  <FileUpload label="Cert. of Completion & MOA" onChange={(e) => handleFileUpload(e, 'certCompletion')} />
                </div>
              </div>
            </SectionBlock>

          </form>

          <aside className="flex flex-col gap-4 rounded-[24px] border border-slate-800 bg-gradient-to-br from-cyan-500/10 via-slate-900 to-violet-500/10 p-4 sm:rounded-[28px] sm:p-6">
            <h3 className="text-lg font-bold text-white sm:text-xl">Report Preview</h3>

            <div className="space-y-3 text-sm text-slate-300 sm:space-y-4">
              <Box label="Student" value={form.studentName} />
              <Box label="Program" value={form.degreeProgram} />
              <Box label="Organization" value={form.background.slice(0, 90) + (form.background.length > 90 ? '...' : '')} />
              <Box label="Status" value={status} />
            </div>

            <div className="mt-auto rounded-2xl border border-slate-700 bg-slate-950/70 p-3 sm:p-4">
              <p className="text-[10px] uppercase tracking-[0.2em] text-slate-400 sm:text-xs">Current flow</p>
              <div className="mt-4 space-y-2.5 sm:space-y-3">
                {sectionOrder.map((section, index) => (
                  <div key={section} className="flex items-center gap-3">
                    <span className="flex h-6 w-6 items-center justify-center rounded-full bg-cyan-500/15 text-[10px] font-bold text-cyan-300 sm:h-7 sm:w-7 sm:text-xs">
                      {index + 1}
                    </span>
                    <span className="text-sm text-slate-200 sm:text-base">{section}</span>
                  </div>
                ))}
              </div>
            </div>

            <button
              type="button"
              onClick={handleGenerate}
              disabled={loading}
              className="mt-2 w-full rounded-2xl bg-gradient-to-r from-cyan-500 to-violet-500 px-4 py-3 text-base font-semibold text-white shadow-glow transition hover:opacity-95 disabled:cursor-not-allowed disabled:opacity-60 sm:mt-4"
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
    <section className="rounded-2xl border border-slate-800 bg-slate-950/60 p-4 sm:rounded-3xl sm:p-5">
      <h2 className="mb-4 text-lg font-bold text-white sm:text-xl">{title}</h2>
      {children}
    </section>
  );
}

function FileUpload({ label, onChange }: { label: string; onChange: (e: React.ChangeEvent<HTMLInputElement>) => void }) {
  return (
    <label className="block text-sm font-medium text-slate-200">
      {label}
      <input type="file" accept="image/png, image/jpeg, image/jpg, image/webp" onChange={onChange} className={fileInputClass} />
    </label>
  );
}

function Card({ label, value, color }: { label: string; value: string; color: 'cyan' | 'violet' | 'emerald' }) {
  const tone = {
    cyan: 'text-cyan-300',
    violet: 'text-violet-300',
    emerald: 'text-emerald-300',
  }[color];

  return (
    <div className="rounded-2xl border border-slate-700 bg-slate-900/80 p-4 sm:p-5">
      <p className="text-[10px] uppercase tracking-[0.2em] text-slate-400 sm:text-sm">{label}</p>
      <p className={`mt-3 text-2xl font-bold sm:mt-4 sm:text-4xl ${tone}`}>{value}</p>
    </div>
  );
}

function Box({ label, value }: { label: string; value: string }) {
  return (
    <div className="rounded-2xl border border-slate-700 bg-slate-950/70 p-3 sm:p-4">
      <p className="text-[10px] uppercase tracking-[0.2em] text-slate-400 sm:text-xs">{label}</p>
      <p className="mt-2 text-sm font-semibold text-white sm:text-base">{value}</p>
    </div>
  );
}