'use client';

import { useState, useEffect } from 'react';
import { Joyride, STATUS, type EventData, type Step } from 'react-joyride';
import { HelpCircle } from 'lucide-react';
import guideScene from '../494356892_23964838003120045_2210637575151932928_n.jpg';

const sectionOrder = [
  'Cover Page',
  'Acknowledgement',
  'Introduction',
  'Organization Analysis',
  'Tasks & Duties',
  'Case Analysis',
  'Reflections',
  'Appendices'
];

const defaultForm = {
  trainingOrganization: 'Zamboanga del Norte Provincial Capitol',
  trainingLocation: 'Dipolog City',
  collegeFaculty: 'College of Computing Studies',
  degreeProgram: 'Bachelor of Science in Information System',
  studentName: 'Ian P. Padilla',
  submittedToName: 'Mr. Erson A. Rodriguez',
  submittedToTitle: 'Associate Dean of the College of Computing Studies',
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
    dailyJournalLayout: 'current',
    reportFooter: {
      preparedByLabel: 'Prepared By:',
      preparedByName: '',
      checkedByLabel: 'Checked By:',
      checkedByName: '',
      checkedByRole: 'Supervising\nStatistical Specialist',
      officeInCharge: '',
      dateLabel: 'Date:',
      dateValue: ''
    },
    dailyJournal: [
      {
        weekNumber: 1,
        totalHours: 40,
        narrative: '',
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
  'mt-2.5 w-full rounded-xl border border-slate-700 bg-slate-950/80 px-3.5 py-3 text-sm text-slate-100 placeholder:text-slate-500 focus:border-cyan-400 focus:outline-none focus:ring-2 focus:ring-cyan-500/30 sm:rounded-2xl sm:px-4 sm:py-3.5 sm:text-base';

const fileInputClass = 
  'mt-2.5 block w-full text-sm text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-cyan-500/10 file:text-cyan-400 hover:file:bg-cyan-500/20';

async function compressImageDataUrl(file: File, maxWidth = 1400, quality = 0.75): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        const scale = Math.min(1, maxWidth / Math.max(img.width, img.height));
        canvas.width = Math.max(1, Math.round(img.width * scale));
        canvas.height = Math.max(1, Math.round(img.height * scale));

        const ctx = canvas.getContext('2d');
        if (!ctx) {
          reject(new Error('Canvas is not available in this browser.'));
          return;
        }

        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL('image/jpeg', quality));
      };
      img.onerror = () => reject(new Error('Failed to process image.'));
      img.src = String(reader.result);
    };
    reader.onerror = () => reject(new Error('Could not read selected image.'));
    reader.readAsDataURL(file);
  });
}

async function readErrorMessage(response: Response): Promise<string> {
  const text = await response.text();
  if (!text) return 'Failed to generate the report';
  try {
    const parsed = JSON.parse(text);
    return parsed?.error || text || 'Failed to generate the report';
  } catch {
    return text || 'Failed to generate the report';
  }
}

function GuideTooltip({ step, tooltipProps, primaryProps }: any) {
  return (
    <div
      {...tooltipProps}
      className="w-full sm:w-max flex flex-col animate-fade-in max-w-[calc(100vw-2rem)] sm:max-w-none"
    >
      {/* Compact Layout: Image + Thought Bubble Side-by-Side on Desktop, Stacked on Mobile */}
      <div className="flex flex-col sm:flex-row gap-2 sm:gap-3 items-start">
        {/* Guide Character Image - Minimized */}
        <div className="relative overflow-hidden rounded-[16px] ring-1 ring-cyan-500/20 flex-shrink-0 mx-auto sm:mx-0">
          <img
            src={guideScene.src}
            alt="Guide illustration"
            className="h-28 w-28 sm:h-32 sm:w-32 object-cover brightness-105 contrast-110"
          />
          <div className="absolute inset-0 bg-gradient-to-t from-black/10 to-transparent" />
        </div>

        {/* Thought Bubble Card - Compact & Adjacent */}
        <div className="relative rounded-[18px] border-2 border-cyan-500/50 bg-gradient-to-br from-slate-950/90 via-slate-900/85 to-slate-800/80 p-3 sm:p-4 shadow-[0_12px_36px_rgba(34,211,238,0.18),inset_0_1px_0_rgba(255,255,255,0.1)] backdrop-blur-xl transition-all duration-300 hover:border-cyan-400/60 hover:shadow-[0_16px_48px_rgba(34,211,238,0.28)] w-full sm:w-[260px]">
          {/* Decorative cloud shapes */}
          <div className="absolute -top-2 -left-2 w-6 h-6 bg-slate-900/40 rounded-full blur-md" />
          <div className="absolute -top-1 left-2 w-4 h-4 bg-slate-900/30 rounded-full blur-md" />
          <div className="absolute -top-2 -right-2 w-5 h-5 bg-slate-900/40 rounded-full blur-md" />

          {/* Bubble pointer tail - pointing to image (hidden on mobile) */}
          <div className="hidden sm:block absolute -left-3 top-6 h-0 w-0 border-l-0 border-r-6 border-t-5 border-b-5 border-r-cyan-500/50 border-t-transparent border-b-transparent" />

          <div className="space-y-2 sm:space-y-2.5 relative z-10">
            <div className="flex items-center justify-between gap-2 flex-wrap">
              <span className="text-[8px] sm:text-[9px] font-bold uppercase tracking-[0.25em] text-cyan-300/90">Guide</span>
              <span className="rounded-full border border-cyan-500/40 bg-cyan-500/15 px-2 py-0.5 text-[8px] sm:text-[9px] font-semibold text-cyan-200/90 backdrop-blur-sm">
                {step.target?.replace('.', '') || 'Tip'}
              </span>
            </div>

            <p className="text-xs sm:text-xs leading-5 text-slate-200/95 font-medium">{step.content}</p>

            <div className="flex items-center justify-between pt-1 gap-2 flex-wrap">
              <button
                type="button"
                className="text-xs font-semibold text-slate-400 hover:text-slate-300 transition-colors duration-200 px-2 py-1.5 rounded hover:bg-slate-800/40"
                {...tooltipProps.closeProps}
              >
                Skip
              </button>

              <button
                type="button"
                {...primaryProps}
                className="rounded-full bg-gradient-to-r from-cyan-500 to-violet-500 px-3 py-1.5 text-xs font-bold text-white shadow-lg shadow-cyan-500/25 transition-all duration-200 hover:shadow-[0_8px_24px_rgba(34,211,238,0.4)] hover:brightness-110 active:scale-95 whitespace-nowrap"
              >
                {primaryProps.title}
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* Global animation style */}
      <style>{`
        @keyframes fadeIn {
          from { opacity: 0; transform: translateY(8px); }
          to { opacity: 1; transform: translateY(0); }
        }
        .animate-fade-in {
          animation: fadeIn 0.4s ease-out;
        }
      `}</style>
    </div>
  );
}

export default function Home() {
  const [form, setForm] = useState(defaultForm);
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState('Draft ready');
  
  // Tour State
  const [runTour, setRunTour] = useState(false);
  const [isMounted, setIsMounted] = useState(false);

  // Load form data from localStorage on mount
  useEffect(() => {
    setIsMounted(true);
    const savedForm = localStorage.getItem('narrativeReportForm');
    if (savedForm) {
      try {
        setForm(JSON.parse(savedForm));
      } catch (error) {
        console.error('Error loading saved form data:', error);
        setForm(defaultForm);
      }
    }
    setRunTour(true);
  }, []);

  // Save form data to localStorage whenever it changes
  useEffect(() => {
    if (isMounted) {
      localStorage.setItem('narrativeReportForm', JSON.stringify(form));
    }
  }, [form, isMounted]);

  const tourSteps: Step[] = [
    {
      target: '.tour-start-btn',
      content: 'Welcome to the JRMSU Narrative Report Generator! Replace the sample values with your own data as you go through each section.',
      skipBeacon: true,
    },
    {
      target: '.tour-form',
      content: 'This is the main form area. Replace the sample text with your real OJT details, from the cover page down to the appendices.',
    },
    {
      target: '.tour-front-page',
      content: 'Start with the Front Page section. Replace the current sample information with your training organization, location, student name, and school details.',
    },
    {
      target: '.tour-acknowledgement',
      content: 'In the Acknowledgement section, replace the example message with your own gratitude and appreciation to the people who helped you.',
    },
    {
      target: '.tour-introduction',
      content: 'Use the Introduction section to replace the sample background with your actual host organization details, including the vision, mission, values, and services.',
    },
    {
      target: '.tour-organization-analysis',
      content: 'The Organization Analysis section should reflect your real evaluation of the company. Replace the sample strengths, weaknesses, opportunities, threats, and recommendations with your own.',
    },
    {
      target: '.tour-tasks-duties',
      content: 'In Tasks and Duties, replace the sample tasks with your actual assigned duties and the procedures you followed during training.',
    },
    {
      target: '.tour-case-analysis',
      content: 'The Case Analysis section should describe your real challenges, actions taken, and lessons learned. Replace all sample entries with your actual experience.',
    },
    {
      target: '.tour-reflections',
      content: 'Use the Reflections section to replace the sample reflection with your honest evaluation and your real connection to your course and goals.',
    },
    {
      target: '.tour-appendices',
      content: 'Finally, update the appendices by replacing the sample journal details, photos, and supporting documents with your actual records.',
    },
    {
      target: '.tour-generate-btn',
      content: 'When all sample content has been replaced with your real information, click Generate Report to create your final DOCX file.',
    }
  ];

  const handleJoyrideCallback = (data: EventData) => {
    const { status, type, step } = data;
    const finishedStatuses = new Set<string>([STATUS.FINISHED, STATUS.SKIPPED]);

    if (finishedStatuses.has(status)) {
      setRunTour(false);
    }
  };

  const updateField = (key: keyof typeof defaultForm, value: string) => {
    setForm((prev) => ({ ...prev, [key]: value }));
  };

  const updateAppendixText = (key: keyof typeof defaultForm.appendices, value: string) => {
    setForm((prev) => ({
      ...prev,
      appendices: { ...prev.appendices, [key]: value }
    }));
  };

  const updateReportFooter = (field: keyof NonNullable<typeof defaultForm.appendices.reportFooter>, value: string) => {
    setForm((prev) => ({
      ...prev,
      appendices: {
        ...prev.appendices,
        reportFooter: {
          ...prev.appendices.reportFooter,
          [field]: value
        }
      }
    }));
  };

  const updateDailyJournalLayout = (layout: 'current' | 'report') => {
    setForm((prev) => ({
      ...prev,
      appendices: { ...prev.appendices, dailyJournalLayout: layout }
    }));
  };

  const addWeek = () => {
    setForm((prev) => {
      const nextWeekNum = prev.appendices.dailyJournal.length + 1;
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
                day: `Day ${i + 1}`,
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

  const addDay = (weekIndex: number) => {
    setForm((prev) => {
      const journal = [...prev.appendices.dailyJournal];
      const week = journal[weekIndex];
      week.activities.push({ day: `Day ${week.activities.length + 1}`, date: '', accomplishment: '', hours: 8 });
      week.totalHours = week.activities.reduce((acc, curr) => acc + (Number(curr.hours) || 0), 0);
      journal[weekIndex].activities = week.activities.map((activity, idx) => ({
        ...activity,
        day: activity.day && activity.day.startsWith('Day ') ? `Day ${idx + 1}` : activity.day
      }));
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
      journal[weekIndex].activities = journal[weekIndex].activities
        .filter((_, idx) => idx !== dayIndex)
        .map((activity, idx) => ({
          ...activity,
          day: activity.day && activity.day.startsWith('Day ') ? `Day ${idx + 1}` : activity.day
        }));
      journal[weekIndex].totalHours = journal[weekIndex].activities.reduce((acc, curr) => acc + (Number(curr.hours) || 0), 0);
      return { ...prev, appendices: { ...prev.appendices, dailyJournal: journal } };
    });
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>, appendixKey: keyof typeof defaultForm.appendices) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (!file.type.startsWith('image/')) {
      alert('Please upload an image file (PNG, JPG).');
      return;
    }
    try {
      const compressed = await compressImageDataUrl(file);
      setForm((prev) => ({
        ...prev,
        appendices: { ...prev.appendices, [appendixKey]: compressed }
      }));
    } catch (error) {
      console.error(error);
      alert('The selected image could not be processed.');
    }
  };

  const handleJournalImageUpload = async (e: React.ChangeEvent<HTMLInputElement>, weekIndex: number) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (!file.type.startsWith('image/')) {
      alert('Please upload an image file (PNG, JPG).');
      return;
    }
    try {
      const compressed = await compressImageDataUrl(file, 1200, 0.7);
      setForm((prev) => {
        const newJournal = [...prev.appendices.dailyJournal];
        const isReportLayout = prev.appendices.dailyJournalLayout === 'report';
        const imageKey = isReportLayout ? 'reportLayoutImages' : 'currentLayoutImages';
        
        // Initialize the layout-specific image array if it doesn't exist
        if (!(newJournal[weekIndex] as any)[imageKey]) {
          (newJournal[weekIndex] as any)[imageKey] = [];
        }
        
        const images = (newJournal[weekIndex] as any)[imageKey];
        if (images.length < 4) {
          images.push({ base64: compressed, detail: `Photo documentation for Week ${newJournal[weekIndex].weekNumber}` });
        } else {
          alert('Maximum of 4 images allowed per week.');
        }
        return { ...prev, appendices: { ...prev.appendices, dailyJournal: newJournal } };
      });
    } catch (error) {
      console.error(error);
      alert('The selected journal image could not be processed.');
    }
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
        const errorMessage = await readErrorMessage(response);
        throw new Error(errorMessage);
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
    <main className="min-h-screen bg-slate-950 px-3 sm:px-4 md:px-6 lg:px-8 py-6 sm:py-8 md:py-12 text-slate-100">
      {isMounted && (
        <Joyride
          steps={tourSteps}
          run={runTour}
          continuous={true}
          onEvent={handleJoyrideCallback}
          tooltipComponent={GuideTooltip}
          options={{
            primaryColor: '#06b6d4',
            backgroundColor: '#0f172a',
            textColor: '#f1f5f9',
            zIndex: 1000,
            overlayColor: 'rgba(15, 23, 42, 0.48)',
            beaconSize: 42,
            arrowColor: '#0f172a',
          }}
          styles={{
            arrow: {
              color: '#0f172a',
            },
            floater: {
              filter: 'drop-shadow(0 18px 40px rgba(34, 211, 238, 0.15))',
            },
            tooltip: {
              borderRadius: 22,
              padding: 0,
              backgroundColor: 'transparent',
              boxShadow: 'none',
            },
            buttonPrimary: {
              backgroundColor: '#06b6d4',
              borderRadius: 999,
              fontWeight: 700,
              color: '#ffffff',
              padding: '0.7rem 1rem',
            },
            buttonBack: {
              color: '#cbd5e1',
            },
            buttonSkip: {
              color: '#cbd5e1',
            },
          }}
          locale={{
            back: 'Back',
            close: 'Close',
            last: 'Finish',
            next: 'Next',
            skip: 'Skip',
          }}
        />
      )}

      <div className="mx-auto max-w-7xl space-y-8 sm:space-y-12">
        <header className="overflow-hidden rounded-[24px] border border-cyan-500/20 bg-slate-900/80 p-6 shadow-glow backdrop-blur-xl sm:rounded-[32px] sm:p-10">
          <div className="flex flex-col gap-6 md:flex-row md:items-center md:justify-between">
            <div className="space-y-2">
              <p className="text-xs uppercase tracking-[0.25em] text-cyan-300 sm:text-sm font-semibold">JRMSU</p>
              <h1 className="text-3xl font-black tracking-tight text-white sm:text-4xl md:text-5xl">
                Narrative Report Generator
              </h1>
              <p className="text-sm text-slate-300">
                Replace the sample details below with your actual OJT information before generating the report.
              </p>
            </div>
            
            <div className="flex flex-col sm:flex-row items-start sm:items-center gap-4">
              <button
                onClick={() => setRunTour(true)}
                className="tour-start-btn flex items-center gap-2 rounded-2xl bg-cyan-500/20 px-4 sm:px-5 py-2.5 sm:py-3.5 text-xs sm:text-sm font-bold text-cyan-300 hover:bg-cyan-500/30 transition-colors border border-cyan-500/30 active:scale-95 touch-auto"
                type="button"
              >
                <HelpCircle size={18} className="sm:w-5 sm:h-5" />
                <span>Show Me How</span>
              </button>

              <div className="rounded-2xl border border-cyan-500/30 bg-cyan-500/10 px-4 py-3 text-left sm:px-5 sm:py-3.5 md:text-right">
                <p className="text-[10px] uppercase tracking-[0.25em] text-cyan-200 sm:text-xs font-semibold">Status</p>
                <p className="mt-1.5 text-lg font-bold text-cyan-300 sm:text-xl">{status}</p>
              </div>
            </div>
          </div>
        </header>

        <section className="grid gap-4 sm:grid-cols-3 sm:gap-6">
          <Card label="Sections" value="8" color="cyan" />
          <Card label="Format" value="DOCX" color="violet" />
          <Card label="Mode" value="Web" color="emerald" />
        </section>

        <div className="grid gap-8 lg:grid-cols-[1.4fr_0.6fr]">
          <form className="tour-form space-y-8 rounded-[24px] border border-slate-800 bg-slate-900/80 p-6 sm:rounded-[32px] sm:p-8">
            
            {/* Added tour-front-page class via modified SectionBlock */}
            <SectionBlock title="Front Page (Cover Page Details)" className="tour-front-page">
              <div className="space-y-5">
                <div className="grid gap-3 sm:gap-5 md:grid-cols-2">
                  <label className="block text-xs sm:text-sm font-medium text-slate-200">
                    <span className="text-cyan-300/70 text-[10px] uppercase tracking-wider block mb-1">Organization</span>
                    <input value={form.trainingOrganization} onChange={(e) => updateField('trainingOrganization', e.target.value)} className={fieldClass} />
                  </label>
                  <label className="block text-xs sm:text-sm font-medium text-slate-200">
                    <span className="text-cyan-300/70 text-[10px] uppercase tracking-wider block mb-1">Location</span>
                    <input value={form.trainingLocation} onChange={(e) => updateField('trainingLocation', e.target.value)} className={fieldClass} />
                  </label>
                </div>
                <div className="grid gap-3 sm:gap-5 md:grid-cols-2">
                  <label className="block text-xs sm:text-sm font-medium text-slate-200">
                    <span className="text-cyan-300/70 text-[10px] uppercase tracking-wider block mb-1">College/Faculty</span>
                    <input value={form.collegeFaculty} onChange={(e) => updateField('collegeFaculty', e.target.value)} className={fieldClass} />
                  </label>
                  <label className="block text-xs sm:text-sm font-medium text-slate-200">
                    <span className="text-cyan-300/70 text-[10px] uppercase tracking-wider block mb-1">Degree Program</span>
                    <input value={form.degreeProgram} onChange={(e) => updateField('degreeProgram', e.target.value)} className={fieldClass} />
                  </label>
                </div>
                <div className="grid gap-3 sm:gap-5 md:grid-cols-3">
                  <label className="block text-xs sm:text-sm font-medium text-slate-200">
                    <span className="text-cyan-300/70 text-[10px] uppercase tracking-wider block mb-1">Student Name</span>
                    <input value={form.studentName} onChange={(e) => updateField('studentName', e.target.value)} className={fieldClass} />
                  </label>
                  <label className="block text-xs sm:text-sm font-medium text-slate-200">
                    <span className="text-cyan-300/70 text-[10px] uppercase tracking-wider block mb-1">Adviser Name</span>
                    <input value={form.submittedToName} onChange={(e) => updateField('submittedToName', e.target.value)} className={fieldClass} />
                  </label>
                  <label className="block text-xs sm:text-sm font-medium text-slate-200">
                    <span className="text-cyan-300/70 text-[10px] uppercase tracking-wider block mb-1">Title/Position</span>
                    <input value={form.submittedToTitle} onChange={(e) => updateField('submittedToTitle', e.target.value)} className={fieldClass} />
                  </label>
                </div>
              </div>
            </SectionBlock>

            <SectionBlock title="1. Acknowledgement" className="tour-acknowledgement">
              <div className="space-y-4 sm:space-y-5">
                <label className="block text-xs sm:text-sm font-medium text-slate-200">
                  <span className="text-cyan-300/70 text-[10px] uppercase tracking-wider block mb-1">Student Name</span>
                  <input value={form.studentName} onChange={(e) => updateField('studentName', e.target.value)} className={fieldClass} />
                </label>
                <label className="block text-xs sm:text-sm font-medium text-slate-200">
                  <span className="text-cyan-300/70 text-[10px] uppercase tracking-wider block mb-1">Degree Program</span>
                  <input value={form.degreeProgram} onChange={(e) => updateField('degreeProgram', e.target.value)} className={fieldClass} />
                </label>
                <label className="block text-xs sm:text-sm font-medium text-slate-200">
                  <span className="text-cyan-300/70 text-[10px] uppercase tracking-wider block mb-1">Acknowledgement</span>
                  <textarea value={form.acknowledgement} onChange={(e) => updateField('acknowledgement', e.target.value)} rows={4} className={fieldClass} />
                </label>
              </div>
            </SectionBlock>

            <SectionBlock title="2. Introduction" className="tour-introduction">
              <div className="space-y-4 sm:space-y-5">
                <label className="block text-xs sm:text-sm font-medium text-slate-200">
                  <span className="text-cyan-300/70 text-[10px] uppercase tracking-wider block mb-1">Background</span>
                  <textarea value={form.background} onChange={(e) => updateField('background', e.target.value)} rows={4} className={fieldClass} />
                </label>
                <div className="grid gap-3 sm:gap-5 md:grid-cols-2">
                  <label className="block text-xs sm:text-sm font-medium text-slate-200">
                    <span className="text-cyan-300/70 text-[10px] uppercase tracking-wider block mb-1">Vision</span>
                    <input value={form.vision} onChange={(e) => updateField('vision', e.target.value)} className={fieldClass} />
                  </label>
                  <label className="block text-xs sm:text-sm font-medium text-slate-200">
                    <span className="text-cyan-300/70 text-[10px] uppercase tracking-wider block mb-1">Mission</span>
                    <input value={form.mission} onChange={(e) => updateField('mission', e.target.value)} className={fieldClass} />
                  </label>
                </div>
                <div className="grid gap-3 sm:gap-5 md:grid-cols-3">
                  <label className="block text-xs sm:text-sm font-medium text-slate-200">
                    <span className="text-emerald-300/70 text-[10px] uppercase tracking-wider block mb-1">Objectives</span>
                    <textarea value={form.objectives} onChange={(e) => updateField('objectives', e.target.value)} rows={3} className={fieldClass} />
                  </label>
                  <label className="block text-xs sm:text-sm font-medium text-slate-200">
                    <span className="text-violet-300/70 text-[10px] uppercase tracking-wider block mb-1">Core Values</span>
                    <textarea value={form.coreValues} onChange={(e) => updateField('coreValues', e.target.value)} rows={3} className={fieldClass} />
                  </label>
                  <label className="block text-xs sm:text-sm font-medium text-slate-200">
                    <span className="text-orange-300/70 text-[10px] uppercase tracking-wider block mb-1">Services</span>
                    <textarea value={form.services} onChange={(e) => updateField('services', e.target.value)} rows={3} className={fieldClass} />
                  </label>
                </div>
              </div>
            </SectionBlock>

            <SectionBlock title="3. Organization Analysis" className="tour-organization-analysis">
              <div className="space-y-5">
                <div className="grid gap-3 sm:gap-5 md:grid-cols-2">
                  <label className="block text-xs sm:text-sm font-medium text-slate-200">
                    <span className="text-emerald-300/70 text-[10px] uppercase tracking-wider block mb-1">✓ Strengths</span>
                    <textarea value={form.strengths} onChange={(e) => updateField('strengths', e.target.value)} rows={4} className={fieldClass} />
                  </label>
                  <label className="block text-xs sm:text-sm font-medium text-slate-200">
                    <span className="text-rose-300/70 text-[10px] uppercase tracking-wider block mb-1">✗ Weaknesses</span>
                    <textarea value={form.weaknesses} onChange={(e) => updateField('weaknesses', e.target.value)} rows={4} className={fieldClass} />
                  </label>
                  <label className="block text-xs sm:text-sm font-medium text-slate-200">
                    <span className="text-cyan-300/70 text-[10px] uppercase tracking-wider block mb-1">◆ Opportunities</span>
                    <textarea value={form.opportunities} onChange={(e) => updateField('opportunities', e.target.value)} rows={4} className={fieldClass} />
                  </label>
                  <label className="block text-xs sm:text-sm font-medium text-slate-200">
                    <span className="text-yellow-300/70 text-[10px] uppercase tracking-wider block mb-1">⚠ Threats</span>
                    <textarea value={form.threats} onChange={(e) => updateField('threats', e.target.value)} rows={4} className={fieldClass} />
                  </label>
                </div>
                <label className="block text-xs sm:text-sm font-medium text-slate-200">
                  <span className="text-blue-300/70 text-[10px] uppercase tracking-wider block mb-1">Recommendations</span>
                  <textarea value={form.recommendations} onChange={(e) => updateField('recommendations', e.target.value)} rows={4} className={fieldClass} />
                </label>
              </div>
            </SectionBlock>

            <SectionBlock title="4. Tasks and Duties" className="tour-tasks-duties">
              <div className="space-y-4 sm:space-y-5">
                <label className="block text-xs sm:text-sm font-medium text-slate-200">
                  <span className="text-emerald-300/70 text-[10px] uppercase tracking-wider block mb-1">Assigned Tasks</span>
                  <textarea value={form.tasks} onChange={(e) => updateField('tasks', e.target.value)} rows={4} className={fieldClass} />
                </label>
                <label className="block text-xs sm:text-sm font-medium text-slate-200">
                  <span className="text-blue-300/70 text-[10px] uppercase tracking-wider block mb-1">Procedures Conformed</span>
                  <textarea value={form.procedures} onChange={(e) => updateField('procedures', e.target.value)} rows={4} className={fieldClass} />
                </label>
              </div>
            </SectionBlock>

            <SectionBlock title="5. Case Analysis" className="tour-case-analysis">
              <div className="space-y-5">
                <div className="grid gap-3 sm:gap-5 md:grid-cols-2">
                  <label className="block text-xs sm:text-sm font-medium text-slate-200">
                    <span className="text-rose-300/70 text-[10px] uppercase tracking-wider block mb-1">Problem 1</span>
                    <textarea value={form.issue1} onChange={(e) => updateField('issue1', e.target.value)} rows={4} className={fieldClass} />
                  </label>
                  <label className="block text-xs sm:text-sm font-medium text-slate-200">
                    <span className="text-emerald-300/70 text-[10px] uppercase tracking-wider block mb-1">Solution 1</span>
                    <textarea value={form.issue1Action} onChange={(e) => updateField('issue1Action', e.target.value)} rows={4} className={fieldClass} />
                  </label>
                  <label className="block text-xs sm:text-sm font-medium text-slate-200">
                    <span className="text-rose-300/70 text-[10px] uppercase tracking-wider block mb-1">Problem 2</span>
                    <textarea value={form.issue2} onChange={(e) => updateField('issue2', e.target.value)} rows={4} className={fieldClass} />
                  </label>
                  <label className="block text-xs sm:text-sm font-medium text-slate-200">
                    <span className="text-emerald-300/70 text-[10px] uppercase tracking-wider block mb-1">Solution 2</span>
                    <textarea value={form.issue2Action} onChange={(e) => updateField('issue2Action', e.target.value)} rows={4} className={fieldClass} />
                  </label>
                </div>
                <label className="block text-xs sm:text-sm font-medium text-slate-200">
                  <span className="text-yellow-300/70 text-[10px] uppercase tracking-wider block mb-1">Lessons Learned</span>
                  <textarea value={form.lessons} onChange={(e) => updateField('lessons', e.target.value)} rows={4} className={fieldClass} />
                </label>
              </div>
            </SectionBlock>

            <SectionBlock title="6. Reflections" className="tour-reflections">
              <div className="space-y-4 sm:space-y-5">
                <label className="block text-xs sm:text-sm font-medium text-slate-200">
                  <span className="text-violet-300/70 text-[10px] uppercase tracking-wider block mb-1">Self-Evaluation</span>
                  <textarea value={form.selfEvaluation} onChange={(e) => updateField('selfEvaluation', e.target.value)} rows={4} className={fieldClass} />
                </label>
                <label className="block text-xs sm:text-sm font-medium text-slate-200">
                  <span className="text-violet-300/70 text-[10px] uppercase tracking-wider block mb-1">Relevancy to Course & Goals</span>
                  <textarea value={form.relevancy} onChange={(e) => updateField('relevancy', e.target.value)} rows={4} className={fieldClass} />
                </label>
              </div>
            </SectionBlock>

            <SectionBlock title="7. Appendices" className="tour-appendices">
              <div className="space-y-6 sm:space-y-8">
                <div className="space-y-5">
                  <div className="flex flex-col gap-3 rounded-2xl border border-slate-800 bg-slate-950/60 p-4 sm:p-4">
                    <div className="flex items-center justify-between gap-2">
                      <h3 className="text-lg font-bold text-cyan-300">Daily Journal</h3>
                      <button type="button" onClick={addWeek} className="rounded-xl bg-cyan-500/20 px-3 sm:px-3.5 py-2.5 sm:py-2 text-xs font-semibold text-cyan-300 hover:bg-cyan-500/30 transition active:scale-95 whitespace-nowrap">
                        + Add Week
                      </button>
                    </div>
                    <div className="flex flex-col gap-2 sm:flex-row sm:items-center">
                      <span className="text-[10px] font-semibold uppercase tracking-[0.2em] text-slate-400">Layout</span>
                      <div className="flex flex-wrap gap-2 sm:gap-3">
                        <label className="flex items-center gap-2 rounded-full border border-slate-700 bg-slate-900 px-3 py-1.5 text-xs text-slate-200">
                          <input type="radio" name="daily-journal-layout" checked={form.appendices.dailyJournalLayout === 'current'} onChange={() => updateDailyJournalLayout('current')} className="accent-cyan-400" />
                          Weekly layout
                        </label>
                        <label className="flex items-center gap-2 rounded-full border border-slate-700 bg-slate-900 px-3 py-1.5 text-xs text-slate-200">
                          <input type="radio" name="daily-journal-layout" checked={form.appendices.dailyJournalLayout === 'report'} onChange={() => updateDailyJournalLayout('report')} className="accent-cyan-400" />
                          Daily layout
                        </label>
                      </div>
                    </div>
                  </div>

                  {form.appendices.dailyJournalLayout === 'report' ? (
                    <div className="space-y-4">
                      {form.appendices.dailyJournal.map((week, weekIdx) => (
                        <div key={weekIdx} className="space-y-4 rounded-2xl border border-slate-800 bg-slate-900/60 p-4 sm:p-5">
                          <div className="flex items-center justify-between gap-2">
                            <span className="font-bold text-white">WEEK {week.weekNumber} • Total Hours: {week.totalHours}</span>
                            {form.appendices.dailyJournal.length > 1 && (
                              <button type="button" onClick={() => removeWeek(weekIdx)} className="text-xs text-rose-400 hover:text-rose-300 px-2 py-1 rounded hover:bg-rose-900/20 transition active:scale-95">
                                Remove Week
                              </button>
                            )}
                          </div>

                          <div className="overflow-hidden rounded-xl border border-slate-800 bg-slate-950/60">
                            <div className="grid grid-cols-1 gap-px bg-slate-800 text-[10px] font-bold uppercase tracking-[0.18em] text-slate-300 sm:grid-cols-[0.9fr_1.8fr_2.2fr]">
                              <div className="bg-slate-900/80 px-3 py-2">Date</div>
                              <div className="bg-slate-900/80 px-3 py-2">Accomplishment</div>
                              <div className="bg-slate-900/80 px-3 py-2">Documentation</div>
                            </div>

                            {week.activities.map((act, dayIdx) => (
                              <div key={dayIdx} className="grid grid-cols-1 gap-3 border-t border-slate-800 bg-slate-950/40 p-3 sm:grid-cols-[0.9fr_1.8fr_2.2fr] sm:gap-0 sm:gap-x-3 sm:p-0">
                                <div className="sm:border-r sm:border-slate-800 sm:p-3">
                                  <input type="text" value={act.date} onChange={(e) => updateDayActivity(weekIdx, dayIdx, 'date', e.target.value)} className="w-full rounded-lg border border-slate-700 bg-slate-900 px-2.5 py-2 text-xs text-slate-100 focus:border-cyan-400 focus:outline-none" placeholder="Date" />
                                </div>
                                <div className="sm:border-r sm:border-slate-800 sm:p-3">
                                  <textarea rows={3} value={act.accomplishment} onChange={(e) => updateDayActivity(weekIdx, dayIdx, 'accomplishment', e.target.value)} className="w-full rounded-lg border border-slate-700 bg-slate-900 px-2.5 py-2 text-xs text-slate-100 focus:border-cyan-400 focus:outline-none" placeholder="Accomplishment" />
                                </div>
                                <div className="sm:p-3">
                                  <div className="space-y-2">
                                    <input type="file" accept="image/png, image/jpeg, image/jpg, image/webp" onChange={(e) => {
                                      const targetIndex = dayIdx;
                                      const file = e.target.files?.[0];
                                      if (!file) return;
                                      if (!file.type.startsWith('image/')) {
                                        alert('Please upload an image file (PNG, JPG).');
                                        return;
                                      }
                                      compressImageDataUrl(file, 1200, 0.7).then((compressed) => {
                                        setForm((prev) => {
                                          const newJournal = [...prev.appendices.dailyJournal];
                                          const imageEntry = { base64: compressed, detail: `Documentation for ${act.day || 'Daily activity'}` };
                                          const isReportLayout = prev.appendices.dailyJournalLayout === 'report';
                                          
                                          // Use layout-specific image storage
                                          const imageKey = isReportLayout ? 'reportLayoutImages' : 'currentLayoutImages';
                                          const existingImages = [...((newJournal[weekIdx] as any)[imageKey] || [])];
                                          const imageIndex = Math.min(targetIndex, existingImages.length);
                                          existingImages[imageIndex] = imageEntry;
                                          (newJournal[weekIdx] as any)[imageKey] = existingImages;
                                          
                                          return { ...prev, appendices: { ...prev.appendices, dailyJournal: newJournal } };
                                        });
                                      }).catch(() => alert('The selected journal image could not be processed.'));
                                    }} className={fileInputClass} />
                                    {(() => {
                                      const isReportLayout = form.appendices.dailyJournalLayout === 'report';
                                      const imageKey = isReportLayout ? 'reportLayoutImages' : 'currentLayoutImages';
                                      const layoutImages = (week as any)[imageKey];
                                      return layoutImages && layoutImages[dayIdx] && (
                                        <div className="rounded-lg border border-slate-800 bg-slate-900/80 p-2">
                                          <img src={layoutImages[dayIdx].base64} alt="Documentation" className="h-20 w-full rounded-md object-cover" />
                                        </div>
                                      );
                                    })()}
                                  </div>
                                </div>
                              </div>
                            ))}
                          </div>

                          <div className="flex items-center justify-between gap-2">
                            <button type="button" onClick={() => addDay(weekIdx)} className="text-xs font-semibold text-cyan-400 hover:text-cyan-300 px-3 py-1.5 rounded hover:bg-cyan-900/20 transition active:scale-95">
                              + Add Day
                            </button>
                            <span className="text-[10px] uppercase tracking-[0.18em] text-slate-400">Daily layout</span>
                          </div>
                        </div>
                      ))}

                      <div className="rounded-2xl border border-slate-800 bg-slate-950/60 p-4">
                        <p className="mb-3 text-[10px] font-semibold uppercase tracking-[0.2em] text-slate-400">Report Footer</p>
                        <div className="grid gap-3 md:grid-cols-2">
                          <label className="block text-xs sm:text-sm font-medium text-slate-200">
                            <span className="text-cyan-300/70 text-[10px] uppercase tracking-wider block mb-1">Prepared By Name</span>
                            <input type="text" value={form.appendices.reportFooter.preparedByName} onChange={(e) => updateReportFooter('preparedByName', e.target.value)} className={fieldClass} />
                          </label>
                          <label className="block text-xs sm:text-sm font-medium text-slate-200">
                            <span className="text-cyan-300/70 text-[10px] uppercase tracking-wider block mb-1">Checked By Name</span>
                            <input type="text" value={form.appendices.reportFooter.checkedByName} onChange={(e) => updateReportFooter('checkedByName', e.target.value)} className={fieldClass} />
                          </label>
                          <label className="block text-xs sm:text-sm font-medium text-slate-200 md:col-span-2">
                            <span className="text-cyan-300/70 text-[10px] uppercase tracking-wider block mb-1">Checked By Role</span>
                            <input type="text" value={form.appendices.reportFooter.checkedByRole} onChange={(e) => updateReportFooter('checkedByRole', e.target.value)} className={fieldClass} />
                          </label>
                          <label className="block text-xs sm:text-sm font-medium text-slate-200">
                            <span className="text-cyan-300/70 text-[10px] uppercase tracking-wider block mb-1">Date Value</span>
                            <input type="text" value={form.appendices.reportFooter.dateValue} onChange={(e) => updateReportFooter('dateValue', e.target.value)} className={fieldClass} />
                          </label>
                          <div className="block text-xs sm:text-sm font-medium text-slate-200 md:col-span-2">
                            <span className="text-cyan-300/70 text-[10px] uppercase tracking-wider block mb-1">Office-in-Charge</span>
                            <div className="px-4 py-2 rounded border border-slate-700 bg-slate-800/50 text-slate-300">Office-in-Charge</div>
                          </div>
                        </div>
                      </div>
                    </div>
                  ) : (
                    <div className="space-y-4">
                      {form.appendices.dailyJournal.map((week, weekIdx) => (
                        <div key={weekIdx} className="space-y-4 rounded-2xl border border-slate-800 bg-slate-900/60 p-5">
                          <div className="flex items-center justify-between">
                            <span className="font-bold text-white">WEEK {week.weekNumber} (Total Hours: {week.totalHours})</span>
                            {form.appendices.dailyJournal.length > 1 && (
                              <button type="button" onClick={() => removeWeek(weekIdx)} className="text-xs text-rose-400 hover:text-rose-300 px-2 py-1 rounded hover:bg-rose-900/20 transition active:scale-95">
                                Remove Week
                              </button>
                            )}
                          </div>
                          <div className="space-y-3">
                            <p className="text-xs font-semibold text-slate-400 uppercase tracking-wider">Daily Activities</p>
                            {week.activities.map((act, dayIdx) => (
                              <div key={dayIdx} className="grid gap-2 sm:gap-2.5 grid-cols-2 sm:grid-cols-[1fr_1fr_2.5fr_1fr_auto] items-end bg-slate-950/60 p-3 rounded-xl border border-slate-800">
                                <div className="col-span-1">
                                  <label className="text-[9px] sm:text-xs text-slate-400 font-semibold uppercase tracking-wider">Day</label>
                                  <input type="text" placeholder="Day" value={act.day} onChange={(e) => updateDayActivity(weekIdx, dayIdx, 'day', e.target.value)} className="w-full rounded-lg border border-slate-700 bg-slate-900 px-2 sm:px-2.5 py-1 sm:py-1.5 text-xs text-slate-100 mt-0.5" />
                                </div>
                                <div className="col-span-1">
                                  <label className="text-[9px] sm:text-xs text-slate-400 font-semibold uppercase tracking-wider">Date</label>
                                  <input type="text" placeholder="Date" value={act.date} onChange={(e) => updateDayActivity(weekIdx, dayIdx, 'date', e.target.value)} className="w-full rounded-lg border border-slate-700 bg-slate-900 px-2 sm:px-2.5 py-1 sm:py-1.5 text-xs text-slate-100 mt-0.5" />
                                </div>
                                <div className="col-span-2 sm:col-span-1">
                                  <label className="text-[9px] sm:text-xs text-slate-400 font-semibold uppercase tracking-wider">Accomplishment</label>
                                  <input type="text" placeholder="Accomplishment" value={act.accomplishment} onChange={(e) => updateDayActivity(weekIdx, dayIdx, 'accomplishment', e.target.value)} className="w-full rounded-lg border border-slate-700 bg-slate-900 px-2 sm:px-2.5 py-1 sm:py-1.5 text-xs text-slate-100 mt-0.5" />
                                </div>
                                <div className="col-span-1">
                                  <label className="text-[9px] sm:text-xs text-slate-400 font-semibold uppercase tracking-wider">Hrs</label>
                                  <input type="number" placeholder="Hrs" value={act.hours} onChange={(e) => updateDayActivity(weekIdx, dayIdx, 'hours', e.target.value)} className="w-full rounded-lg border border-slate-700 bg-slate-900 px-2 sm:px-2.5 py-1 sm:py-1.5 text-xs text-slate-100 mt-0.5" />
                                </div>
                                <button type="button" onClick={() => removeDay(weekIdx, dayIdx)} className="col-span-1 text-rose-400 text-xs px-2 py-1.5 hover:text-rose-300 rounded hover:bg-rose-900/20">✕</button>
                              </div>
                            ))}
                            <button type="button" onClick={() => addDay(weekIdx)} className="mt-2 text-xs font-semibold text-cyan-400 hover:text-cyan-300 px-3 py-1.5 rounded hover:bg-cyan-900/20 transition active:scale-95">
                              + Add Day
                            </button>
                          </div>
                          <label className="block text-sm font-medium text-slate-200 pt-2">
                            Weekly Narrative Report
                            <textarea value={week.narrative} onChange={(e) => {
                                const newJournal = [...form.appendices.dailyJournal];
                                newJournal[weekIdx].narrative = e.target.value;
                                setForm((prev) => ({ ...prev, appendices: { ...prev.appendices, dailyJournal: newJournal } }));
                              }} rows={2} className={fieldClass} />
                          </label>
                          <div className="pt-2">
                            <label className="block text-sm font-medium text-slate-200">
                              Upload Pictures for Week {week.weekNumber} (Max 4)
                              <input type="file" accept="image/png, image/jpeg, image/jpg, image/webp" onChange={(e) => handleJournalImageUpload(e, weekIdx)} className={fileInputClass} />
                            </label>
                            {(() => {
                              const isReportLayout = form.appendices.dailyJournalLayout === 'report';
                              const imageKey = isReportLayout ? 'reportLayoutImages' : 'currentLayoutImages';
                              const layoutImages = (week as any)[imageKey];
                              const imageCount = layoutImages?.length || 0;
                              return layoutImages && layoutImages.length > 0 && (
                                <div className={`mt-4 ${imageCount === 3 ? 'flex flex-wrap justify-center gap-4' : 'grid grid-cols-2 gap-4'}`}>
                                  {layoutImages.map((img: any, imgIdx: number) => (
                                    <div key={imgIdx} className={`flex flex-col bg-slate-950 p-3 rounded-xl border border-slate-800 hover:border-slate-700 transition ${imageCount === 3 && imgIdx === 2 ? 'w-full sm:w-1/2' : ''}`}>
                                      <div className="relative mb-2 overflow-hidden rounded-lg bg-slate-900 h-32 sm:h-40 flex items-center justify-center">
                                        <img src={img.base64} alt="Preview" className="w-full h-full object-cover" />
                                        <button type="button" onClick={() => {
                                            const newJournal = [...form.appendices.dailyJournal];
                                            (newJournal[weekIdx] as any)[imageKey] = ((newJournal[weekIdx] as any)[imageKey] as any[]).filter((_: any, idx: number) => idx !== imgIdx);
                                            setForm((prev) => ({ ...prev, appendices: { ...prev.appendices, dailyJournal: newJournal } }));
                                          }} className="absolute top-1 right-1 text-xs text-rose-400 bg-slate-900/80 px-2 py-1 rounded hover:bg-rose-900/60 transition">✕</button>
                                      </div>
                                      <input type="text" placeholder="Caption / Picture Detail" value={img.detail} onChange={(e) => {
                                          const newJournal = [...form.appendices.dailyJournal];
                                          ((newJournal[weekIdx] as any)[imageKey])[imgIdx].detail = e.target.value;
                                          setForm((prev) => ({ ...prev, appendices: { ...prev.appendices, dailyJournal: newJournal } }));
                                        }} className="w-full rounded-lg border border-slate-700 bg-slate-900 px-2.5 py-1.5 text-xs text-slate-100 focus:border-cyan-400 focus:outline-none" />
                                    </div>
                                  ))}
                                </div>
                              );
                            })()}
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
                <label className="block text-xs sm:text-sm font-medium text-slate-200">
                  <span className="text-cyan-300/70 text-[10px] uppercase tracking-wider block mb-1">PRIME Report</span>
                  <textarea value={form.appendices.primeNarrative} onChange={(e) => updateAppendixText('primeNarrative', e.target.value)} rows={3} className={fieldClass} />
                </label>
                <div className="grid gap-3 sm:gap-5 md:grid-cols-2 pt-2">
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

          <aside className="flex flex-col gap-5 rounded-[24px] border border-slate-800 bg-gradient-to-br from-cyan-500/10 via-slate-900 to-violet-500/10 p-6 sm:rounded-[32px] sm:p-8">
            <h3 className="text-lg font-bold text-white sm:text-xl">Report Preview</h3>
            <div className="space-y-4 text-sm text-slate-300">
              <Box label="Student" value={form.studentName} />
              <Box label="Program" value={form.degreeProgram} />
              <Box label="Organization" value={form.trainingOrganization} />
              <Box label="Status" value={status} />
            </div>
            <div className="mt-auto rounded-2xl border border-slate-700 bg-slate-950/70 p-4 sm:p-5">
              <p className="text-[10px] uppercase tracking-[0.2em] text-slate-400 sm:text-xs font-semibold">Current flow</p>
              <div className="mt-4 space-y-3 sm:space-y-3.5">
                {sectionOrder.map((section, index) => (
                  <div key={section} className="flex items-center gap-3.5">
                    <span className="flex h-7 w-7 items-center justify-center rounded-full bg-cyan-500/15 text-xs font-bold text-cyan-300">{index + 1}</span>
                    <span className="text-sm text-slate-200 sm:text-base font-medium">{section}</span>
                  </div>
                ))}
              </div>
            </div>
            <button
              type="button"
              onClick={handleGenerate}
              disabled={loading}
              className="tour-generate-btn mt-2 w-full rounded-2xl bg-gradient-to-r from-cyan-500 to-violet-500 px-5 py-3.5 sm:py-4 text-base font-semibold text-white shadow-glow transition hover:opacity-95 disabled:cursor-not-allowed disabled:opacity-60 active:scale-95"
            >
              {loading ? 'Generating...' : 'Generate Report'}
            </button>
          </aside>
        </div>

        <footer className="flex justify-center pb-2">
          <a href="https://ianpadilla-opal.vercel.app/" target="_blank" rel="noopener noreferrer" className="text-sm text-slate-400 transition hover:text-cyan-300">
            Developed by <span className="font-semibold text-cyan-300">Ian P. Padilla</span>
          </a>
        </footer>
      </div>
    </main>
  );
}

// Updated to accept className for the tour targeting
function SectionBlock({ title, children, className = '' }: { title: string; children: React.ReactNode; className?: string }) {
  return (
    <section className={`rounded-2xl border border-slate-800 bg-slate-950/60 p-5 sm:rounded-3xl sm:p-6 space-y-5 ${className}`}>
      <h2 className="text-lg font-bold text-white sm:text-xl">{title}</h2>
      {children}
    </section>
  );
}

function FileUpload({ label, onChange }: { label: string; onChange: (e: React.ChangeEvent<HTMLInputElement>) => void }) {
  return (
    <label className="block text-sm font-medium text-slate-200 space-y-1.5">
      {label}
      <input type="file" accept="image/png, image/jpeg, image/jpg, image/webp" onChange={onChange} className={fileInputClass} />
    </label>
  );
}

function Card({ label, value, color }: { label: string; value: string; color: 'cyan' | 'violet' | 'emerald' }) {
  const tone = { cyan: 'text-cyan-300', violet: 'text-violet-300', emerald: 'text-emerald-300' }[color];
  return (
    <div className="rounded-2xl border border-slate-700 bg-slate-900/80 p-5 sm:p-6 space-y-2">
      <p className="text-[10px] uppercase tracking-[0.2em] text-slate-400 sm:text-xs font-semibold">{label}</p>
      <p className={`text-2xl font-bold sm:text-4xl ${tone}`}>{value}</p>
    </div>
  );
}

function Box({ label, value }: { label: string; value: string }) {
  return (
    <div className="rounded-2xl border border-slate-700 bg-slate-950/70 p-4 space-y-1.5">
      <p className="text-[10px] uppercase tracking-[0.2em] text-slate-400 sm:text-xs font-semibold">{label}</p>
      <p className="text-sm font-semibold text-white sm:text-base">{value}</p>
    </div>
  );
}