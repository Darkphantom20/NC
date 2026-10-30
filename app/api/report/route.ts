import fs from 'fs';
import path from 'path';
import JSZip from 'jszip';
import { createClient } from '@supabase/supabase-js';
import { Document, Packer, Paragraph, TextRun, AlignmentType, ImageRun, Header, Footer, Table, TableRow, TableCell, WidthType, BorderStyle, TableLayoutType } from 'docx';

export const runtime = 'nodejs';

const OUTPUT_FILE_NAME = process.env.OUTPUT_FILE_NAME || 'JRMSU_Narrative_Report.docx';
const HEADER_IMAGE_PATH = path.join(process.cwd(), 'public', 'Picture1.png');
const HEADER_IMAGE = fs.existsSync(HEADER_IMAGE_PATH) ? fs.readFileSync(HEADER_IMAGE_PATH) : null;
const SCANNED_TEMPLATE_HEADER_PATH = path.join(process.cwd(), 'public', 'Untitled.png');
const SCANNED_TEMPLATE_HEADER = fs.existsSync(SCANNED_TEMPLATE_HEADER_PATH) ? fs.readFileSync(SCANNED_TEMPLATE_HEADER_PATH) : null;
const HD_TEMPLATE_HEADER_PATH = path.join(process.cwd(), 'public', 'hd.png');
const HD_TEMPLATE_HEADER = fs.existsSync(HD_TEMPLATE_HEADER_PATH) ? fs.readFileSync(HD_TEMPLATE_HEADER_PATH) : null;

interface AppendixImage {
  base64: string;
  detail?: string;
}

interface DailyActivity {
  day: string;
  date: string;
  accomplishment: string;
  hours: number | string;
}

interface DailyJournalWeek {
  weekNumber: number;
  activities: DailyActivity[];
  totalHours: number | string;
  narrative?: string;
  images?: (string | AppendixImage)[]; // For backwards compatibility
  reportLayoutImages?: (string | AppendixImage)[]; // Report layout only
  currentLayoutImages?: (string | AppendixImage)[]; // Current layout only
}

interface ReportFooterData {
  preparedByLabel?: string;
  preparedByName?: string;
  checkedByLabel?: string;
  checkedByName?: string;
  checkedByRole?: string;
  officeInCharge?: string;
  dateLabel?: string;
  dateValue?: string;
}

interface AppendicesData {
  dailyJournalLayout?: 'current' | 'report' | 'no-photo';
  reportFooter?: ReportFooterData;
  dailyJournal?: DailyJournalWeek[];
  certParticipation?: string;
  primeNarrative?: string;
  primeReportImages?: (string | AppendixImage)[];
  resume?: string;
  grades?: string;
  medicalWaiver?: string;
  letterAcceptance?: string;
  dtr?: string;
  ratingSheet?: string;
  certCompletion?: string;
}

interface SectionData {
  title: string;
  content: any;
  isBullet?: boolean;
}

const SCANNED_APPENDIX_DEFINITIONS = [
  { letter: 'A', title: 'Journal', imageKey: null },
  { letter: 'B', title: 'Certificate of Participation of Prime Seminar', imageKey: 'certParticipation' },
  { letter: 'C', title: 'One-Page Narrative Report of PRIME', imageKey: null },
  { letter: 'D', title: 'Resume and Application Letter', imageKey: 'resume' },
  { letter: 'E', title: 'Evaluation of Grades & Validated Enrolment Form', imageKey: 'grades' },
  { letter: 'G', title: 'Medical Certificate & Parent’s Waiver', imageKey: 'medicalWaiver' },
  { letter: 'I', title: 'Letter of Acceptance', imageKey: 'letterAcceptance' },
  { letter: 'J', title: 'Daily Time Record (Number of Hours Rendered)', imageKey: 'dtr' },
  { letter: 'K', title: 'Student’s Intern Rating Sheet & Performance Evaluation', imageKey: 'ratingSheet' },
  { letter: 'M', title: 'Certificate of Completion & Memorandum of Agreement', imageKey: 'certCompletion' },
] as const;

type ScannedAppendixLetter = typeof SCANNED_APPENDIX_DEFINITIONS[number]['letter'];

function getScannedJournalImageCandidates(
  week: DailyJournalWeek,
  layout: AppendicesData['dailyJournalLayout'],
): ((string | AppendixImage) | null | undefined)[] {
  if (layout === 'no-photo') return [];
  return layout === 'report'
    ? week.reportLayoutImages || week.images || []
    : week.currentLayoutImages || week.images || [];
}

function getScannedJournalImageSources(week: DailyJournalWeek, layout: AppendicesData['dailyJournalLayout']): (string | AppendixImage)[] {
  return getScannedJournalImageCandidates(week, layout).filter((image): image is string | AppendixImage => Boolean(
    typeof image === 'string' ? image.trim() : image?.base64?.trim(),
  ));
}

function getScannedJournalWeeks(data: any): DailyJournalWeek[] {
  if (!Array.isArray(data.appendices?.dailyJournal)) return [];
  return data.appendices.dailyJournal.filter((week: DailyJournalWeek) => {
    const activities = Array.isArray(week.activities) ? week.activities : [];
    const hasActivityData = activities.some((activity) =>
      Boolean(
        String(activity.date || '').trim()
        || String(activity.accomplishment || '').trim()
        || (Number(activity.hours) || 0) > 0,
      ),
    );
    const hasPhoto = getScannedJournalImageSources(week, data.appendices?.dailyJournalLayout).length > 0;
    const hasCurrentLayoutNarrative = data.appendices?.dailyJournalLayout !== 'report'
      && data.appendices?.dailyJournalLayout !== 'no-photo'
      && Boolean(String(week.narrative || '').trim());
    return hasActivityData || hasPhoto || hasCurrentLayoutNarrative;
  });
}

function getScannedJournalCaption(layout: AppendicesData['dailyJournalLayout']): string {
  if (layout === 'report') return 'Table 1: Daily Accomplishment Journal';
  if (layout === 'no-photo') return 'Table 1: Daily Journal (No Pictures)';
  return 'Table 1: Weekly Journal';
}

function getScannedAppendixImages(data: any, letter: ScannedAppendixLetter): (string | AppendixImage)[] {
  const definition = SCANNED_APPENDIX_DEFINITIONS.find((entry) => entry.letter === letter);
  if (!definition) return [];
  const value = letter === 'C'
    ? data.appendices?.primeReportImages
    : definition.imageKey ? data.appendices?.[definition.imageKey] : null;
  const candidates = Array.isArray(value) ? value : value ? [value] : [];
  return candidates.filter((image): image is string | AppendixImage => Boolean(
    typeof image === 'string' ? image.trim() : image && typeof image.base64 === 'string' && image.base64.trim(),
  ));
}

function getScannedAppendixAvailability(data: any): { letter: ScannedAppendixLetter; title: string; hasContent: boolean }[] {
  return SCANNED_APPENDIX_DEFINITIONS.map((entry) => {
    let hasContent = false;
    if (entry.letter === 'A') {
      hasContent = getScannedJournalWeeks(data).length > 0;
    } else if (entry.letter === 'C') {
      const narrative = typeof data.appendices?.primeNarrative === 'string'
        ? data.appendices.primeNarrative.trim()
        : '';
      const images = getScannedAppendixImages(data, entry.letter);
      hasContent = Boolean(narrative) || images.some((image) => Boolean(getScannedImageData(image, `Appendix ${entry.letter}`)));
    } else {
      const images = getScannedAppendixImages(data, entry.letter);
      hasContent = images.some((image) => Boolean(getScannedImageData(image, `Appendix ${entry.letter}`)));
    }
    return { letter: entry.letter, title: entry.title, hasContent };
  });
}

function buildHeaderImage(headerImage: Buffer | null = HEADER_IMAGE, width = 624, height = 125) {
  if (!headerImage) return [];
  return [
    new ImageRun({
      data: headerImage,
      type: 'png',
      transformation: { width, height },
    }),
  ];
}

const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;
const hasSupabaseConfig = Boolean(supabaseUrl && supabaseKey);

function ensureArray(val: any): string[] {
  if (Array.isArray(val)) return val.map(v => String(v).trim()).filter(Boolean);
  if (typeof val === 'string' && val.trim().length > 0) {
    return val.split(/\r?\n/).map(v => v.trim()).filter(Boolean);
  }
  return [];
}

function splitTemplateItems(value: any, count = 3): string[] {
  const items = ensureArray(value);
  const padded = [...items];
  while (padded.length < count) padded.push('');
  return padded.slice(0, count);
}

function escapeXml(value: any): string {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function decodeXmlText(value: string): string {
  return value
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function buildScannedTableOfContents(data: any): { title: string; page: string; level: number }[] {
  const secondPlacement = Boolean(data.schoolPlacement?.enabled);
  const org1 = data.trainingOrganization || 'TRAINING ORGANIZATION';
  const org2 = data.schoolPlacement?.trainingOrganization || 'SCHOOL TRAINING ORGANIZATION';
  const appendicesStart = secondPlacement ? 11 : 6;

  const entries: { title: string; page: string; level: number }[] = [
    { title: 'Acknowledgement', page: 'i', level: 0 },
    { title: 'Table of Contents', page: 'ii', level: 0 },
    { title: 'List of Tables / Figures', page: 'iv', level: 0 },
    { title: `I. ${org1.toUpperCase()}`, page: '1', level: 0 },
    { title: 'Introduction', page: '1', level: 1 },
    { title: 'Background of the Organization', page: '1', level: 2 },
    { title: 'Vision, Mission, and Objectives', page: '1', level: 2 },
    { title: 'Organizational Structure', page: '1', level: 2 },
    { title: 'The Products and Services Offered', page: '1', level: 2 },
    { title: 'Organization / Company Analysis', page: '2', level: 1 },
    { title: 'Strengths and Weaknesses of the Organization', page: '2', level: 2 },
    { title: 'Opportunities and Threats of the Organization', page: '2', level: 2 },
    { title: 'Recommendations for Improvement', page: '2', level: 2 },
    { title: 'Tasks and Duties', page: '3', level: 1 },
    { title: 'Tasks and Responsibilities Assigned', page: '3', level: 2 },
    { title: 'Duties and Procedures Conformed', page: '3', level: 2 },
    { title: 'Case Analysis', page: '4', level: 1 },
    { title: 'Identified Issues / Problems', page: '4', level: 2 },
    { title: 'Strategies Undertaken', page: '4', level: 2 },
    { title: 'Lessons Learned', page: '4', level: 2 },
    { title: 'Reflections', page: '5', level: 1 },
    { title: 'Self-Evaluation', page: '5', level: 2 },
    { title: 'Relevance to Course and Goals', page: '5', level: 2 },
  ];

  if (secondPlacement) {
    entries.push(
      { title: `II. ${org2.toUpperCase()}`, page: '6', level: 0 },
      { title: 'Introduction', page: '6', level: 1 },
      { title: 'Background of the Organization', page: '6', level: 2 },
      { title: 'Vision, Mission, and Objectives', page: '6', level: 2 },
      { title: 'Organizational Structure', page: '6', level: 2 },
      { title: 'The Products and Services Offered', page: '6', level: 2 },
      { title: 'Organization / Company Analysis', page: '7', level: 1 },
      { title: 'Strengths and Weaknesses of the Organization', page: '7', level: 2 },
      { title: 'Opportunities and Threats of the Organization', page: '7', level: 2 },
      { title: 'Recommendations for Improvement', page: '7', level: 2 },
      { title: 'Tasks and Duties', page: '8', level: 1 },
      { title: 'Tasks and Responsibilities Assigned', page: '8', level: 2 },
      { title: 'Duties and Procedures Conformed', page: '8', level: 2 },
      { title: 'Case Analysis', page: '9', level: 1 },
      { title: 'Identified Issues / Problems', page: '9', level: 2 },
      { title: 'Strategies Undertaken', page: '9', level: 2 },
      { title: 'Lessons Learned', page: '9', level: 2 },
      { title: 'Reflections', page: '10', level: 1 },
      { title: 'Self-Evaluation', page: '10', level: 2 },
      { title: 'Relevance to Course and Goals', page: '10', level: 2 },
    );
  }

  const availableAppendices = getScannedAppendixAvailability(data).filter((entry) => entry.hasContent);
  if (availableAppendices.length > 0) {
    entries.push({ title: 'Appendices', page: String(appendicesStart), level: 0 });
    availableAppendices.forEach((entry, index) => {
      entries.push({
        title: `Appendix ${entry.letter}: ${entry.title}`,
        page: String(appendicesStart + 1 + index),
        level: 1,
      });
    });
  }

  return entries;
}

function replaceScannedTableOfContents(xml: string, data: any): string {
  const entries = buildScannedTableOfContents(data);
  const generatedParagraphs = entries.map((entry) => {
    const styleLevel = entry.level + 1;
    return `<w:p><w:pPr><w:pStyle w:val="TOC${styleLevel}"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9360"/></w:tabs></w:pPr><w:r><w:t xml:space="preserve">${escapeXml(entry.title)}</w:t></w:r><w:r><w:tab/><w:t xml:space="preserve">${escapeXml(entry.page)}</w:t></w:r></w:p>`;
  }).join('');

  const paragraphTexts = Array.from(xml.matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g), ([paragraph]) => {
    let text = '';
    paragraph.replace(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g, (_match, encodedText) => {
      text += decodeXmlText(encodedText);
      return _match;
    });
    return text.trim();
  });
  let contentsHeadingIndex = -1;
  let listHeadingIndex = -1;
  for (let index = 0; index < paragraphTexts.length; index += 1) {
    if (contentsHeadingIndex < 0 && /^Table of Contents$/i.test(paragraphTexts[index])) {
      contentsHeadingIndex = index;
      continue;
    }
    if (contentsHeadingIndex >= 0 && /^List of Tables\s*\/\s*Figures$/i.test(paragraphTexts[index])) {
      listHeadingIndex = index;
      break;
    }
  }
  if (contentsHeadingIndex < 0 || listHeadingIndex < 0) return xml;

  let inTableOfContents = false;
  let sawContentsHeading = false;

  return xml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (paragraph) => {
    let text = '';
    paragraph.replace(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g, (_match, encodedText) => {
      text += decodeXmlText(encodedText);
      return _match;
    });

    const trimmed = text.trim();
    if (/^Table of Contents$/i.test(trimmed) && !sawContentsHeading) {
      sawContentsHeading = true;
      inTableOfContents = true;
      return paragraph + generatedParagraphs;
    }
    if (!inTableOfContents) return paragraph;
    if (/^List of Tables\s*\/\s*Figures$/i.test(trimmed)) {
      inTableOfContents = false;
      return paragraph;
    }
    return '';
  });
}

function replaceScannedJournalList(xml: string, data: any): string {
  const weeks = getScannedJournalWeeks(data);
  const generatedEntries = weeks.length > 0
    ? `<w:p><w:pPr><w:pStyle w:val="Caption"/><w:keepNext/></w:pPr><w:r><w:t xml:space="preserve">${escapeXml(getScannedJournalCaption(data.appendices?.dailyJournalLayout))}</w:t></w:r></w:p>`
    : '';
  const paragraphs = Array.from(xml.matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g), (match) => ({
    paragraph: match[0],
    text: getScannedParagraphText(match[0]),
  }));
  const tablesHeadingIndex = paragraphs.findIndex((entry) => /^List of Tables$/i.test(entry.text));
  if (tablesHeadingIndex < 0) return xml;
  const figuresHeadingIndex = paragraphs.findIndex((entry, index) =>
    index > tablesHeadingIndex && /^List of Figures$/i.test(entry.text),
  );
  if (figuresHeadingIndex < 0) return xml;

  let currentIndex = -1;
  let insideList = false;
  return xml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (paragraph) => {
    const text = getScannedParagraphText(paragraph);
    if (/^List of Tables$/i.test(text) && !insideList) {
      insideList = true;
      currentIndex = tablesHeadingIndex;
      return paragraph + generatedEntries;
    }
    if (insideList && /^List of Figures$/i.test(text)) {
      insideList = false;
      return paragraph;
    }
    if (!insideList) return paragraph;
    currentIndex += 1;
    return currentIndex < figuresHeadingIndex ? '' : paragraph;
  });
}

function replaceScannedFigureList(xml: string, data: any): string {
  const entries: { title: string; page: string }[] = [];
  const appendicesStart = data.schoolPlacement?.enabled ? 11 : 6;
  const availableAppendices = getScannedAppendixAvailability(data).filter((entry) => entry.hasContent);
  if (data.organizationStructureImage) {
    entries.push({
      title: 'Figure 1: Organizational Chart of ' + (data.trainingOrganization || 'Training Organization'),
      page: '1',
    });
  }
  if (data.schoolPlacement?.enabled && data.schoolPlacement.organizationStructureImage) {
    entries.push({
      title: 'Figure 2: Organizational Chart of ' + (data.schoolPlacement.trainingOrganization || 'School Placement'),
      page: '6',
    });
  }
  if (getScannedJournalWeeks(data).some((week) => getScannedJournalImageSources(week, data.appendices?.dailyJournalLayout).length > 0)) {
    entries.push({
      title: 'Figure 3: Weekly documentation photos',
      page: String(appendicesStart + 1),
    });
  }
  const primeReportPageIndex = availableAppendices.findIndex((entry) => entry.letter === 'C');
  if (primeReportPageIndex >= 0 && getScannedAppendixImages(data, 'C').length > 0) {
    entries.push({
      title: 'Figure 4: PRIME report photos',
      page: String(appendicesStart + 1 + primeReportPageIndex),
    });
  }
  const generatedEntries = entries.map((entry) =>
    `<w:p><w:pPr><w:pStyle w:val="TableofFigures"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9360"/></w:tabs></w:pPr><w:r><w:t xml:space="preserve">${escapeXml(entry.title)}</w:t></w:r><w:r><w:tab/><w:t xml:space="preserve">${escapeXml(entry.page)}</w:t></w:r></w:p>`,
  ).join('');
  let inserted = false;
  return xml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (paragraph) => {
    const text = getScannedParagraphText(paragraph);
    if (/^List of Figures$/i.test(text) && !inserted) {
      inserted = true;
      return paragraph + generatedEntries;
    }
    if (/<w:pStyle\b[^>]*w:val="TableofFigures"/.test(paragraph)) return '';
    return paragraph;
  });
}

function removeScannedSchoolPlacementPages(xml: string, enabled: boolean): string {
  if (enabled) return xml;
  const headings = Array.from(xml.matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g), (match) => ({
    start: match.index ?? 0,
    paragraph: match[0],
    text: getScannedParagraphText(match[0]),
  })).filter((entry) => /<w:pStyle\b[^>]*w:val="Heading1"/.test(entry.paragraph));
  const placementHeading = headings.find((entry) => /^II\.\s/.test(entry.text));
  const appendicesHeading = headings.find((entry) => entry.start > (placementHeading?.start ?? -1) && /^Appendices$/i.test(entry.text));
  if (!placementHeading || !appendicesHeading) return xml;
  return xml.slice(0, placementHeading.start) + xml.slice(appendicesHeading.start);
}

function replaceTextAcrossRuns(xml: string, replacements: Record<string, string | string[]>): string {
  const repeatedReplacementIndexes: Record<string, number> = {};
  return xml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (paragraph) => {
    let text = '';
    paragraph.replace(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g, (match, encodedText, offset) => {
      const decodedText = decodeXmlText(encodedText);
      text += decodedText;
      return match;
    });

    if (!text.trim() && text.length === 0) return paragraph;

    let updatedText = text;
    for (const [placeholder, value] of Object.entries(replacements)) {
      if (Array.isArray(value)) {
        let searchFrom = 0;
        while (searchFrom <= updatedText.length) {
          const foundAt = updatedText.indexOf(placeholder, searchFrom);
          if (foundAt < 0) break;
          const repeatedIndex = repeatedReplacementIndexes[placeholder] ?? 0;
          const replacement = repeatedIndex < value.length ? value[repeatedIndex] : '';
          updatedText =
            updatedText.slice(0, foundAt)
            + String(replacement ?? '')
            + updatedText.slice(foundAt + placeholder.length);
          searchFrom = foundAt + String(replacement ?? '').length;
          repeatedReplacementIndexes[placeholder] = repeatedIndex + 1;
        }
      } else {
        updatedText = updatedText.split(placeholder).join(String(value ?? ''));
      }
    }

    if (updatedText === text) return paragraph;

    let firstTextNode = true;
    return paragraph.replace(/<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>/g, (match) => {
      if (firstTextNode) {
        firstTextNode = false;
        return match.replace(/>[\s\S]*?</, `>${escapeXml(updatedText)}<`);
      }
      return match.replace(/>[\s\S]*?</, '><');
    });
  });
}

function sizeScannedHeaderImage(xml: string): string {
  const headerWidthEmu = 5943600;
  const headerHeightEmu = 1038000;
  let sized = xml
    .replace(/<wp:extent\b[^>]*>/g, `<wp:extent cx="${headerWidthEmu}" cy="${headerHeightEmu}"/>`)
    .replace(/<a:xfrm>\s*<a:off[^>]*\/>\s*<a:ext\b[^>]*\/>\s*<\/a:xfrm>/g,
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${headerWidthEmu}" cy="${headerHeightEmu}"/></a:xfrm>`);
  return sized.replace(/<w:pPr>([\s\S]*?)<\/w:pPr>/g, () =>
    '<w:pPr>' +
      '<w:jc w:val="center"/>' +
      '<w:spacing w:after="60"/>' +
    '</w:pPr>',
  );
}

function getScannedImageData(
  image: string | AppendixImage,
  context: string,
): { data: Buffer; extension: string; width: number; height: number } {
  const dataUrl = typeof image === 'string' ? image : image.base64;
  if (typeof dataUrl !== 'string' || !dataUrl.trim()) {
    throw new Error(`The image for ${context} is invalid.`);
  }
  const match = dataUrl.match(/^data:image\/(jpeg|jpg|png);base64,([\s\S]+)$/i);
  if (!match) {
    throw new Error(`The image for ${context} must be a PNG or JPEG image.`);
  }

  const extension = match[1].toLowerCase() === 'png' ? 'png' : 'jpg';
  const data = Buffer.from(match[2], 'base64');
  if (data.length === 0) {
    throw new Error(`The image for ${context} is empty.`);
  }

  let width = 0;
  let height = 0;
  if (extension === 'png' && data.length >= 24) {
    width = data.readUInt32BE(16);
    height = data.readUInt32BE(20);
  } else if (extension === 'jpg' && data.length >= 4) {
    let offset = 2;
    while (offset < data.length) {
      if (data[offset] !== 0xff) {
        offset += 1;
        continue;
      }

      const marker = data[offset + 1];
      offset += 2;
      if (marker === undefined || marker === 0xd9 || marker === 0xda) break;
      if (marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd7)) continue;
      if (offset + 7 > data.length) break;

      const segmentLength = data.readUInt16BE(offset);
      if ([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(marker)) {
        height = data.readUInt16BE(offset + 3);
        width = data.readUInt16BE(offset + 5);
        break;
      }
      if (segmentLength < 2) break;
      offset += segmentLength;
    }
  }

  if (!width || !height) {
    throw new Error(`The image for ${context} could not be read.`);
  }

  return { data, extension, width, height };
}

function buildScannedJournalImageRun(
  image: { width: number; height: number },
  relationshipId: string,
  drawingId: number,
  name: string,
  maxWidth = 1.45 * 914400,
  maxHeight = 0.9 * 914400,
): string {
  const scale = Math.min(maxWidth / image.width, maxHeight / image.height);
  const width = Math.max(1, Math.round(image.width * scale));
  const height = Math.max(1, Math.round(image.height * scale));

  return `<w:r><w:drawing xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${width}" cy="${height}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${drawingId}" name="${escapeXml(name)}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${drawingId}" name="${escapeXml(name)}"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relationshipId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}

function replaceScannedOrganizationCharts(
  xml: string,
  relationshipsXml: string,
  data: any,
  existingPaths: string[],
): { xml: string; relationshipsXml: string; images: { path: string; data: Buffer }[] } {
  const charts: (string | AppendixImage | null)[] = [
    data.organizationStructureImage || null,
    data.schoolPlacement?.enabled ? data.schoolPlacement.organizationStructureImage || null : null,
  ];
  const existingImagePaths = new Set(existingPaths);
  const relationshipIds = Array.from(relationshipsXml.matchAll(/\bId="rId(\d+)"/g), (match) => Number(match[1]));
  const drawingIds = Array.from(xml.matchAll(/<wp:docPr\b[^>]*\bid="(\d+)"/g), (match) => Number(match[1]));
  let nextRelationshipId = Math.max(0, ...relationshipIds) + 1;
  let nextDrawingId = Math.max(0, ...drawingIds) + 1;
  const images: { path: string; data: Buffer }[] = [];
  const newRelationships: string[] = [];
  let chartIndex = 0;
  let updatedXml = xml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (paragraph) => {
    const text = getScannedParagraphText(paragraph);
    if (/^Figure [12]: Organizational Chart of/.test(text)
      && /<w:pStyle\b[^>]*w:val="Caption"/.test(paragraph)
      && !charts[Number(text.match(/^Figure ([12])/)?.[1]) - 1]) {
      return '';
    }
    if (text !== '[Insert organizational chart here]') return paragraph;
    const chartNumber = chartIndex++;
    const source = charts[chartNumber];
    if (!source) return '';

    const image = getScannedImageData(source, `organization ${chartNumber + 1} chart`);
    const basePath = `word/media/organization-chart-${chartNumber + 1}`;
    let imagePath = `${basePath}.${image.extension}`;
    let suffix = 1;
    while (existingImagePaths.has(imagePath)) {
      imagePath = `${basePath}-${suffix}.${image.extension}`;
      suffix += 1;
    }
    existingImagePaths.add(imagePath);
    const relationshipId = `rId${nextRelationshipId++}`;
    newRelationships.push(`<Relationship Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${imagePath.replace('word/', '')}"/>`);
    images.push({ path: imagePath, data: image.data });
    const drawing = buildScannedJournalImageRun(
      image,
      relationshipId,
      nextDrawingId++,
      `Organization ${chartNumber + 1} chart`,
      6.5 * 914400,
      5.5 * 914400,
    );
    return `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${drawing}</w:p>`;
  });

  if (newRelationships.length > 0) {
    if (!relationshipsXml.includes('</Relationships>')) {
      throw new Error('The scanned-template image relationships are invalid.');
    }
    relationshipsXml = relationshipsXml.replace('</Relationships>', `${newRelationships.join('')}</Relationships>`);
  }
  return { xml: updatedXml, relationshipsXml, images };
}

function getScannedParagraphText(paragraph: string): string {
  let text = '';
  paragraph.replace(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g, (_match, encodedText) => {
    text += decodeXmlText(encodedText);
    return _match;
  });
  return text.trim();
}

function scannedAppendixTextParagraph(text: string, centered = false): string {
  return `<w:p><w:pPr><w:jc w:val="${centered ? 'center' : 'both'}"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/><w:color w:val="000000"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

function formatScannedAppendixParagraphs(xml: string): string {
  return xml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (paragraph) => {
    let formatted = paragraph.replace(/<w:pPr>([\s\S]*?)<\/w:pPr>/, (_match, properties) => {
      const spacing = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>';
      const updatedProperties = /<w:spacing\b[^>]*\/>/.test(properties)
        ? properties.replace(/<w:spacing\b[^>]*\/>/, spacing)
        : `${spacing}${properties}`;
      return `<w:pPr>${updatedProperties}</w:pPr>`;
    });
    if (!/<w:pPr>/.test(formatted)) {
      formatted = formatted.replace(/<w:p\b[^>]*>/, (openingTag) =>
        `${openingTag}<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>`,
      );
    }

    return formatted.replace(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g, (run) => {
      const properties = run.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/)?.[1] || '';
      const boldProperties = properties.match(/<w:b(?:\s[^>]*)?\/?>/g)?.join('') || '';
      const runProperties =
        '<w:rPr>' +
          '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>' +
          '<w:color w:val="000000"/><w:sz w:val="24"/><w:szCs w:val="24"/>' +
          boldProperties +
        '</w:rPr>';
      return /<w:rPr>[\s\S]*?<\/w:rPr>/.test(run)
        ? run.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/, runProperties)
        : run.replace(/<w:r\b[^>]*>/, (openingTag) => `${openingTag}${runProperties}`);
    });
  });
}

function formatScannedTemplateDetails(xml: string): string {
  return xml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (paragraph) => {
    if (!/<w:pStyle\b[^>]*w:val="Body"/.test(paragraph) || !getScannedParagraphText(paragraph)) {
      return paragraph;
    }
    return formatScannedAppendixParagraphs(paragraph);
  });
}

function replaceScannedAppendixSections(
  xml: string,
  relationshipsXml: string,
  data: any,
  existingPaths: string[],
): { xml: string; relationshipsXml: string; images: { path: string; data: Buffer }[] } {
  const availability = getScannedAppendixAvailability(data);
  const availableLetters = new Set(availability.filter((entry) => entry.hasContent).map((entry) => entry.letter));
  const headings = Array.from(xml.matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g), (match) => ({
    start: match.index ?? 0,
    paragraph: match[0],
    text: getScannedParagraphText(match[0]),
  })).filter((paragraph) => (
    /^Appendix [A-N]:/.test(paragraph.text)
    && /<w:pStyle\b[^>]*w:val="Heading2"/.test(paragraph.paragraph)
  ));

  const relationshipIds = Array.from(relationshipsXml.matchAll(/\bId="rId(\d+)"/g), (match) => Number(match[1]));
  let nextRelationshipId = Math.max(0, ...relationshipIds) + 1;
  const drawingIds = Array.from(xml.matchAll(/<wp:docPr\b[^>]*\bid="(\d+)"/g), (match) => Number(match[1]));
  let nextDrawingId = Math.max(0, ...drawingIds) + 1;
  const existingImagePaths = new Set(existingPaths);
  const newRelationships: string[] = [];
  const images: { path: string; data: Buffer }[] = [];

  const addAppendixImage = (image: string | AppendixImage, letter: ScannedAppendixLetter, index: number): string => {
    const imageData = getScannedImageData(image, `Appendix ${letter}`);
    const basePath = `word/media/scanned-appendix-${letter}-${index + 1}`;
    let imagePath = `${basePath}.${imageData.extension}`;
    let suffix = 1;
    while (existingImagePaths.has(imagePath)) {
      imagePath = `${basePath}-${suffix}.${imageData.extension}`;
      suffix += 1;
    }
    existingImagePaths.add(imagePath);

    const relationshipId = `rId${nextRelationshipId}`;
    nextRelationshipId += 1;
    newRelationships.push(`<Relationship Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${imagePath.replace('word/', '')}"/>`);
    images.push({ path: imagePath, data: imageData.data });

    const drawing = buildScannedJournalImageRun(
      imageData,
      relationshipId,
      nextDrawingId,
      `Appendix ${letter} image ${index + 1}`,
      6.5 * 914400,
      9 * 914400,
    );
    nextDrawingId += 1;
    const imageParagraph = `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="120" w:line="240" w:lineRule="auto"/></w:pPr>${drawing}</w:p>`;
    const detail = typeof image === 'string' ? '' : image.detail?.trim() || '';
    return imageParagraph + (detail ? scannedAppendixTextParagraph(detail, true) : '');
  };

  let updatedXml = xml;
  const bodyEnd = xml.lastIndexOf('</w:body>');
  const sectionPropertiesStart = xml.lastIndexOf('<w:sectPr', bodyEnd);
  const finalAppendixEnd = sectionPropertiesStart >= 0 ? sectionPropertiesStart : bodyEnd;

  for (let index = headings.length - 1; index >= 0; index -= 1) {
    const heading = headings[index];
    const letter = heading.text.match(/^Appendix ([A-N]):/)?.[1] as ScannedAppendixLetter;
    const start = heading.start;
    const end = headings[index + 1]?.start ?? finalAppendixEnd;
    if (!availableLetters.has(letter)) {
      updatedXml = updatedXml.slice(0, start) + updatedXml.slice(end);
      continue;
    }

    let section = updatedXml.slice(start, end);
    if (letter !== 'A') {
      const bodyParts: string[] = [];
      if (letter === 'C') {
        const narrative = typeof data.appendices?.primeNarrative === 'string'
          ? data.appendices.primeNarrative.trim()
          : '';
        bodyParts.push(...narrative.split(/\r?\n/).map((line: string) => line.trim()).filter(Boolean).map((line: string) => scannedAppendixTextParagraph(line)));
      }
      getScannedAppendixImages(data, letter).forEach((image, imageIndex) => {
        bodyParts.push(addAppendixImage(image, letter, imageIndex));
      });

      let bodyReplaced = false;
      section = section.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (paragraph) => {
        if (bodyReplaced || !/^\[Insert or scan\b/i.test(getScannedParagraphText(paragraph))) return paragraph;
        bodyReplaced = true;
        return bodyParts.join('');
      });
      if (!bodyReplaced) {
        throw new Error(`The scanned-template content placeholder for Appendix ${letter} is missing.`);
      }
    }
    section = formatScannedAppendixParagraphs(section);
    updatedXml = updatedXml.slice(0, start) + section + updatedXml.slice(end);
  }

  const listParagraphs = Array.from(updatedXml.matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g), (match) => ({
    start: match.index ?? 0,
    paragraph: match[0],
    text: getScannedParagraphText(match[0]),
  }));
  const appendixListHeadingIndex = listParagraphs.findIndex((entry, index) =>
    entry.text === 'Appendices' && /^A\.\s/.test(listParagraphs[index + 1]?.text || ''),
  );
  if (appendixListHeadingIndex >= 0) {
    const listItems = [];
    for (let index = appendixListHeadingIndex + 1; index < listParagraphs.length && /^[A-N]\.\s/.test(listParagraphs[index].text); index += 1) {
      listItems.push(listParagraphs[index]);
    }
    const listIsEmpty = availableLetters.size === 0;
    const removals = [
      ...(listIsEmpty ? [listParagraphs[appendixListHeadingIndex]] : []),
      ...listItems.filter((item) => !availableLetters.has(item.text[0] as ScannedAppendixLetter)),
    ];
    for (const item of removals.reverse()) {
      updatedXml = updatedXml.slice(0, item.start) + updatedXml.slice(item.start + item.paragraph.length);
    }
  }

  const remainingListParagraphs = Array.from(updatedXml.matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g), (match) => ({
    start: match.index ?? 0,
    paragraph: match[0],
    text: getScannedParagraphText(match[0]),
  }));
  const remainingListHeadingIndex = remainingListParagraphs.findIndex((entry, index) =>
    entry.text === 'Appendices' && /^[A-N]\.\s/.test(remainingListParagraphs[index + 1]?.text || ''),
  );
  if (remainingListHeadingIndex >= 0) {
    let listEndIndex = remainingListHeadingIndex + 1;
    while (listEndIndex + 1 < remainingListParagraphs.length && /^[A-N]\.\s/.test(remainingListParagraphs[listEndIndex + 1].text)) {
      listEndIndex += 1;
    }
    const listStart = remainingListParagraphs[remainingListHeadingIndex].start;
    const lastListParagraph = remainingListParagraphs[listEndIndex];
    const listEnd = lastListParagraph.start + lastListParagraph.paragraph.length;
    updatedXml = updatedXml.slice(0, listStart)
      + formatScannedAppendixParagraphs(updatedXml.slice(listStart, listEnd))
      + updatedXml.slice(listEnd);
  }

  if (newRelationships.length > 0) {
    if (!relationshipsXml.includes('</Relationships>')) {
      throw new Error('The scanned-template image relationships are invalid.');
    }
    relationshipsXml = relationshipsXml.replace('</Relationships>', `${newRelationships.join('')}</Relationships>`);
  }

  return { xml: updatedXml, relationshipsXml, images };
}

function replaceScannedJournalTable(
  xml: string,
  relationshipsXml: string,
  data: any,
): { xml: string; relationshipsXml: string; images: { path: string; data: Buffer }[] } {
  const weeks = getScannedJournalWeeks(data);
  const images: { path: string; data: Buffer }[] = [];
  const newRelationships: string[] = [];
  const relationshipIds = Array.from(relationshipsXml.matchAll(/\bId="rId(\d+)"/g), (match) => Number(match[1]));
  let nextRelationshipId = Math.max(0, ...relationshipIds) + 1;
  let drawingId = 1000;
  const layout = data.appendices?.dailyJournalLayout === 'report'
    ? 'report'
    : data.appendices?.dailyJournalLayout === 'no-photo' ? 'no-photo' : 'current';
  const selectedColumns = layout === 'report' ? [1, 2, 4, 3] : layout === 'no-photo' ? [0, 1, 2, 4] : [0, 1, 2, 3, 4];
  const columnWidths = layout === 'report'
    ? [1500, 4200, 1000, 2660]
    : layout === 'no-photo' ? [1300, 1500, 4900, 1660] : [];

  const replaceCellText = (cell: string, value: string) => {
    let firstTextNode = true;
    return cell.replace(/<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>/g, () => {
      if (!firstTextNode) return '';
      firstTextNode = false;
      return `<w:t xml:space="preserve">${escapeXml(value)}</w:t>`;
    });
  };

  const replaceCellImage = (cell: string, imageParagraphs: string) => cell.replace(
    /<w:p\b[^>]*>[\s\S]*?<\/w:p>/,
    imageParagraphs,
  );
  const setCellWidth = (cell: string, width: number, gridSpan?: number) => {
    let updated = cell.replace(/<w:tcW\b[^>]*\/>/, `<w:tcW w:type="dxa" w:w="${width}"/>`);
    if (gridSpan !== undefined) {
      updated = updated.replace(/<w:gridSpan\b[^>]*\/>/, `<w:gridSpan w:val="${gridSpan}"/>`);
    }
    return updated;
  };
  const projectRowCells = (row: string, widths = columnWidths) => {
    const cells = Array.from(row.matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g), (match) => match[0]);
    return selectedColumns.map((columnIndex, index) => (
      widths.length ? setCellWidth(cells[columnIndex], widths[index]) : cells[columnIndex]
    ));
  };

  const updatedXml = xml.replace(/<w:tbl\b[\s\S]*?<\/w:tbl>/g, (table) => {
    const rows = Array.from(table.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g), (match) => match[0]);
    const rowCells = rows.map((row) => Array.from(row.matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g), (match) => match[0]));
    const headerIndex = rowCells.findIndex((cells) => (
      cells.length === 5
      && cells[0].includes('>Week<')
      && cells[1].includes('>Date<')
      && cells[2].includes('Activities / Tasks Accomplished')
      && cells[3].includes('>Image<')
      && cells[4].includes('>Hours<')
    ));
    if (headerIndex < 0) return table;

    const dataRowIndex = rowCells.findIndex((cells) => cells.some((cell) => cell.includes('[Date range]')));
    const totalRowIndex = rowCells.findIndex((cells) => cells.some((cell) => cell.includes('TOTAL HOURS')));
    if (dataRowIndex < 0 || totalRowIndex < 0 || rowCells[dataRowIndex].length !== 5) {
      throw new Error('The scanned-template journal table has an unexpected layout.');
    }

    const templateDataCells = rowCells[dataRowIndex];
    const populatedRows: string[] = [];
    weeks.forEach((week, weekIndex) => {
      const activities = Array.isArray(week.activities) ? week.activities : [];
      const weekNumber = week.weekNumber || weekIndex + 1;
      const candidates = getScannedJournalImageCandidates(week, layout);
      const sources = getScannedJournalImageSources(week, layout);
      const buildImageParagraphs = (selectedImages: (string | AppendixImage)[], nameSuffix: string) => selectedImages
        .map((source, imageIndex) => {
          const image = getScannedImageData(source, `journal week ${weekNumber} image ${imageIndex + 1}`);
          const relationshipId = `rId${nextRelationshipId++}`;
          const imagePath = `word/media/journal-week-${weekIndex + 1}-${nameSuffix}-${imageIndex + 1}.${image.extension}`;
          const imageRun = buildScannedJournalImageRun(image, relationshipId, drawingId++, `Journal week ${weekNumber} image ${imageIndex + 1}`);
          images.push({ path: imagePath, data: image.data });
          newRelationships.push(`<Relationship Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${imagePath.split('/').pop()}"/>`);
          return `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${imageRun}</w:p>`;
        })
        .join('');
      const buildRow = (
        rowValues: { week: string; date: string; activity: string; hours: string; images: (string | AppendixImage)[] },
        imageSuffix: string,
      ) => {
        const cells = [...templateDataCells];
        cells[0] = replaceCellText(cells[0], rowValues.week);
        cells[1] = replaceCellText(cells[1], rowValues.date);
        cells[2] = replaceCellText(cells[2], rowValues.activity);
        cells[4] = replaceCellText(cells[4], rowValues.hours);
        const imageParagraphs = buildImageParagraphs(rowValues.images, imageSuffix);
        cells[3] = imageParagraphs
          ? replaceCellImage(cells[3], imageParagraphs)
          : replaceCellText(cells[3], '');
        const selectedCells = selectedColumns.map((columnIndex, index) => (
          columnWidths.length ? setCellWidth(cells[columnIndex], columnWidths[index]) : cells[columnIndex]
        ));
        return rows[dataRowIndex].replace(/<w:tc\b[\s\S]*?<\/w:tc>/g, () => selectedCells.shift() || '');
      };

      if (layout === 'current') {
        const dates = activities.map((activity) => String(activity.date || '').trim()).filter(Boolean);
        const dateRange = dates.length > 1 && dates[0] !== dates[dates.length - 1]
          ? `${dates[0]} - ${dates[dates.length - 1]}`
          : dates[0] || '';
        const activitySummary = activities
          .map((activity) => [String(activity.day || '').trim(), String(activity.accomplishment || '').trim()].filter(Boolean).join(': '))
          .filter(Boolean)
          .join('; ');
        const narrative = String(week.narrative || '').trim();
        const journalSummary = [activitySummary, narrative].filter(Boolean).join('; ');
        const weeklyHours = activities.reduce((total, activity) => total + (Number(activity.hours) || 0), 0);
        populatedRows.push(buildRow({
          week: `Week ${weekNumber}`,
          date: dateRange,
          activity: journalSummary,
          hours: String(Number(weeklyHours.toFixed(2))),
          images: sources,
        }, 'weekly'));
        return;
      }

      const rowCount = layout === 'report' ? Math.max(activities.length, candidates.length) : activities.length;
      for (let dayIndex = 0; dayIndex < rowCount; dayIndex += 1) {
        const activity = activities[dayIndex];
        const selectedImage = candidates[dayIndex];
        const rowHasData = Boolean(
          activity
          && (String(activity.date || '').trim()
            || String(activity.accomplishment || '').trim()
            || (Number(activity.hours) || 0) > 0),
        );
        const image = selectedImage && (typeof selectedImage === 'string'
          ? selectedImage.trim()
          : selectedImage.base64.trim())
          ? [selectedImage as string | AppendixImage]
          : [];
        if (!rowHasData && image.length === 0) continue;
        populatedRows.push(buildRow({
          week: layout === 'report'
            ? `${String(activity?.day || `Day ${dayIndex + 1}`).trim()}${activity?.date ? ` - ${String(activity.date).trim()}` : ''}`
            : `Week ${weekNumber} - ${String(activity?.day || `Day ${dayIndex + 1}`).trim()}`,
          date: layout === 'report'
            ? `Week ${weekNumber} - ${String(activity?.day || `Day ${dayIndex + 1}`).trim()}${activity?.date ? `, ${String(activity.date).trim()}` : ''}`
            : '',
          activity: String(activity?.accomplishment || '').trim(),
          hours: String(Number(Number(activity?.hours || 0).toFixed(2))),
          images: image,
        }, `day-${dayIndex + 1}`));
      }
    });

    const totalHours = weeks.reduce((total, week) => {
      const activities = Array.isArray(week.activities) ? week.activities : [];
      return total + activities.reduce((weeklyTotal, activity) => weeklyTotal + (Number(activity.hours) || 0), 0);
    }, 0);
    const totalCells = [...rowCells[totalRowIndex]];
    totalCells[totalCells.length - 1] = replaceCellText(
      totalCells[totalCells.length - 1],
      String(Number(totalHours.toFixed(2))),
    );
    const projectedTotalCells = layout === 'current'
      ? totalCells
      : [
        setCellWidth(
          totalCells[0].replace(/<w:gridSpan\b[^>]*\/>/, `<w:gridSpan w:val="${selectedColumns.length - 1}"/>`),
          columnWidths.slice(0, -1).reduce((total, width) => total + width, 0),
        ),
        setCellWidth(totalCells[totalCells.length - 1], columnWidths[columnWidths.length - 1], 1),
      ];
    const totalRow = rows[totalRowIndex].replace(/<w:tc\b[\s\S]*?<\/w:tc>/g, () => projectedTotalCells.shift() || '');
    const replacementRows = [
      (() => {
        const headerCells = projectRowCells(rows[headerIndex]);
        const labels = layout === 'report'
          ? ['Date', 'Accomplishment', 'Hours', 'Documentation']
          : layout === 'no-photo' ? ['Day', 'Date', 'Accomplishment', 'Hours'] : [];
        labels.forEach((label, index) => {
          headerCells[index] = replaceCellText(headerCells[index], label);
        });
        return rows[headerIndex].replace(/<w:tc\b[\s\S]*?<\/w:tc>/g, () => headerCells.shift() || '');
      })(),
      ...populatedRows,
      totalRow,
    ].join('');
    let rowIndex = -1;
    const updatedTable = table.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, () => {
      rowIndex += 1;
      return rowIndex === 0 ? replacementRows : '';
    });
    if (layout === 'current') return updatedTable;
    const gridColumns = columnWidths.map((width) => `<w:gridCol w:w="${width}"/>`).join('');
    return updatedTable.replace(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/, `<w:tblGrid>${gridColumns}</w:tblGrid>`);
  });
  const captionUpdatedXml = updatedXml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (paragraph) => (
    /<w:pStyle\b[^>]*w:val="Caption"/.test(paragraph)
    && getScannedParagraphText(paragraph) === 'Table 1: Journal Week 1'
      ? replaceCellText(paragraph, getScannedJournalCaption(layout))
      : paragraph
  ));

  if (newRelationships.length > 0) {
    if (!relationshipsXml.includes('</Relationships>')) {
      throw new Error('The scanned-template image relationships are invalid.');
    }
    relationshipsXml = relationshipsXml.replace('</Relationships>', `${newRelationships.join('')}</Relationships>`);
  }

  return { xml: captionUpdatedXml, relationshipsXml, images };
}

async function fillReferenceTemplate(data: any): Promise<Buffer> {
  const templatePath = path.join(process.cwd(), 'public', 'narrative-report-template.docx');
  const templateBuffer = fs.readFileSync(templatePath);
  const zip = await JSZip.loadAsync(templateBuffer);
  const documentRelationshipsPath = 'word/_rels/document.xml.rels';
  const documentRelationshipsFile = zip.file(documentRelationshipsPath);
  if (!documentRelationshipsFile) {
    throw new Error('The scanned-template document relationships are missing.');
  }
  let documentRelationshipsXml = await documentRelationshipsFile.async('string');
  const journalImages: { path: string; data: Buffer }[] = [];

  const schoolPlacement = data.schoolPlacement?.enabled ? data.schoolPlacement : null;
  const perPlacement = (primary: any, secondary: any) => (
    schoolPlacement ? [primary ?? '', secondary ?? ''] : [primary ?? '']
  );
  const perPlacementItem = (primary: any, secondary: any, index: number) => perPlacement(
    splitTemplateItems(primary)[index],
    splitTemplateItems(secondary)[index],
  );
  const trainingDuration = data.trainingDuration || '[Month DD - Month DD, YYYY]  ([000] HRS)';
  const trainingYearMatch = String(trainingDuration).match(/(?:19|20)\d{2}/);
  const trainingMonthYear = trainingYearMatch ? trainingYearMatch[0] : String(new Date().getFullYear());

  const replacements: Record<string, string | string[]> = {
    '[ORGANIZATION 1 NAME]': data.trainingOrganization || '',
    '[ORGANIZATION 2 NAME]': data.schoolPlacement?.enabled ? (data.schoolPlacement.trainingOrganization || '') : '',
    '[Program] – [Year Level]': `${data.degreeProgram || ''}${data.studentYear ? ` - ${data.studentYear}` : ''}`,
    '[Degree Program] – [Year Level]': `${data.degreeProgram || ''}${data.studentYear ? ` - ${data.studentYear}` : ''}`,
    '[Organization 1]': data.trainingOrganization || '',
    '[Organization 2]': data.schoolPlacement?.enabled ? (data.schoolPlacement.trainingOrganization || '') : '',
    'Evaluation of Grades from the Registrar': 'Evaluation of Grades & Validated Enrolment Form',
    'Medical Certificate': 'Medical Certificate & Parent’s Waiver',
    'Student’s Intern Rating Sheet': 'Student’s Intern Rating Sheet & Performance Evaluation',
    'Certificate of Completion': 'Certificate of Completion & Memorandum of Agreement',
    '[Address / Location of Organization 1] and [Address / Location of Organization 2]':
      [data.trainingLocation, data.schoolPlacement?.enabled ? data.schoolPlacement.trainingLocation : ''].filter(Boolean).join(' and '),
    '[Address / Location of Organization 1]': data.trainingLocation || '',
    '[Address / Location of Organization 2]': data.schoolPlacement?.enabled ? (data.schoolPlacement.trainingLocation || '') : '',
    '[College Name]': data.collegeFaculty || '',
    '[Degree Program]': data.degreeProgram || '',
    '[Month DD - Month DD, YYYY]  ([000] HRS)': trainingDuration,
    '[Month DD – Month DD, YYYY]  ([000] HRS)': trainingDuration,
    '[Month DD — Month DD, YYYY]  ([000] HRS)': trainingDuration,
    '[Month YYYY]': trainingMonthYear,
    '[STUDENT FULL NAME]': data.studentName || '',
    '[Program]': data.degreeProgram || '',
    '[Year Level]': data.studentYear || '',
    '[OJT COORDINATOR NAME]': data.submittedToName || '',
    '[Express your gratitude to the host organizations, supervisors, the faculty OJT coordinator, and everyone who supported your training.]': data.acknowledgement || '',
    '[Add a second paragraph to thank your family, classmates, and fellow trainees.]': '',
    '[Student Name]': data.studentName || '',
    '[Describe the history, location, and nature of the organization in two to three paragraphs.]': perPlacement(data.background, schoolPlacement?.background),
    '[Insert the organization’s vision statement.]': perPlacement(data.vision, schoolPlacement?.vision),
    '[Insert the organization’s mission statement.]': perPlacement(data.mission, schoolPlacement?.mission),
    '[Objective 1]': perPlacementItem(data.objectives, schoolPlacement?.objectives, 0),
    '[Objective 2]': perPlacementItem(data.objectives, schoolPlacement?.objectives, 1),
    '[Objective 3]': perPlacementItem(data.objectives, schoolPlacement?.objectives, 2),
    '[Value 1]': perPlacementItem(data.coreValues, schoolPlacement?.coreValues, 0),
    '[Value 2]': perPlacementItem(data.coreValues, schoolPlacement?.coreValues, 1),
    '[Value 3]': perPlacementItem(data.coreValues, schoolPlacement?.coreValues, 2),
    '[Product or service 1]': perPlacementItem(data.services, schoolPlacement?.services, 0),
    '[Product or service 2]': perPlacementItem(data.services, schoolPlacement?.services, 1),
    '[Product or service 3]': perPlacementItem(data.services, schoolPlacement?.services, 2),
    '[Strength 1]': perPlacementItem(data.strengths, schoolPlacement?.strengths, 0),
    '[Strength 2]': perPlacementItem(data.strengths, schoolPlacement?.strengths, 1),
    '[Strength 3]': perPlacementItem(data.strengths, schoolPlacement?.strengths, 2),
    '[Weakness 1]': perPlacementItem(data.weaknesses, schoolPlacement?.weaknesses, 0),
    '[Weakness 2]': perPlacementItem(data.weaknesses, schoolPlacement?.weaknesses, 1),
    '[Weakness 3]': perPlacementItem(data.weaknesses, schoolPlacement?.weaknesses, 2),
    '[Opportunity 1]': perPlacementItem(data.opportunities, schoolPlacement?.opportunities, 0),
    '[Opportunity 2]': perPlacementItem(data.opportunities, schoolPlacement?.opportunities, 1),
    '[Opportunity 3]': perPlacementItem(data.opportunities, schoolPlacement?.opportunities, 2),
    '[Threat 1]': perPlacementItem(data.threats, schoolPlacement?.threats, 0),
    '[Threat 2]': perPlacementItem(data.threats, schoolPlacement?.threats, 1),
    '[Threat 3]': perPlacementItem(data.threats, schoolPlacement?.threats, 2),
    '[Recommendation 1]': perPlacementItem(data.recommendations, schoolPlacement?.recommendations, 0),
    '[Recommendation 2]': perPlacementItem(data.recommendations, schoolPlacement?.recommendations, 1),
    '[Recommendation 3]': perPlacementItem(data.recommendations, schoolPlacement?.recommendations, 2),
    '[Describe the tasks you were assigned during the training period.]': perPlacement(data.tasks, schoolPlacement?.tasks),
    '[Briefly describe the hierarchy and where your assigned unit or office belongs.]': '',
    '[Attach your weekly journal entries. Add one table per week using the format below.]': '',
    '[Duty or procedure 1]': perPlacementItem(data.procedures, schoolPlacement?.procedures, 0),
    '[Duty or procedure 2]': perPlacementItem(data.procedures, schoolPlacement?.procedures, 1),
    '[Duty or procedure 3]': perPlacementItem(data.procedures, schoolPlacement?.procedures, 2),
    '[Describe the issue you encountered.]': [
      data.issue1 || '',
      data.issue2 || '',
      ...(schoolPlacement ? [schoolPlacement.issue1 || '', schoolPlacement.issue2 || ''] : []),
    ],
    '[Explain what was done to address the issue.]': [
      data.issue1Action || '',
      data.issue2Action || '',
      ...(schoolPlacement ? [schoolPlacement.issue1Action || '', schoolPlacement.issue2Action || ''] : []),
    ],
    '[Summarize what you learned from these situations.]': perPlacement(data.lessons, schoolPlacement?.lessons),
    '[Reflect on your growth, strengths, and areas to improve.]': perPlacement(data.selfEvaluation, schoolPlacement?.selfEvaluation),
    '[Explain how the training relates to your degree program and career goals.]': perPlacement(data.relevancy, schoolPlacement?.relevancy),
  };

  for (const name of Object.keys(zip.files)) {
    if (!/^word\/(?:document|header\d+|footer\d+)\.xml$/.test(name)) continue;
    let xml = await zip.file(name)?.async('string');
    if (!xml) continue;
    xml = replaceTextAcrossRuns(xml, replacements);
    if (name === 'word/document.xml') {
      const organizationCharts = replaceScannedOrganizationCharts(xml, documentRelationshipsXml, data, Object.keys(zip.files));
      xml = organizationCharts.xml;
      documentRelationshipsXml = organizationCharts.relationshipsXml;
      journalImages.push(...organizationCharts.images);
      xml = removeScannedSchoolPlacementPages(xml, Boolean(schoolPlacement));
      const journalTable = replaceScannedJournalTable(xml, documentRelationshipsXml, data);
      xml = journalTable.xml;
      documentRelationshipsXml = journalTable.relationshipsXml;
      journalImages.push(...journalTable.images);
      xml = replaceScannedTableOfContents(xml, data);
      xml = replaceScannedJournalList(xml, data);
      xml = replaceScannedFigureList(xml, data);
      const appendixSections = replaceScannedAppendixSections(xml, documentRelationshipsXml, data, Object.keys(zip.files));
      xml = appendixSections.xml;
      documentRelationshipsXml = appendixSections.relationshipsXml;
      journalImages.push(...appendixSections.images);
      xml = formatScannedTemplateDetails(xml);
    }
    if (/^word\/header\d+\.xml$/.test(name)) {
      xml = sizeScannedHeaderImage(xml);
    }
    zip.file(name, xml);
  }

  zip.file(documentRelationshipsPath, documentRelationshipsXml);

  const scannedHeader = fs.existsSync(HD_TEMPLATE_HEADER_PATH)
    ? fs.readFileSync(HD_TEMPLATE_HEADER_PATH)
    : null;
  if (scannedHeader) {
    const headerRelationshipFiles = Object.keys(zip.files).filter((name) => /^word\/_rels\/header\d+\.xml\.rels$/.test(name));
    for (const relationshipFile of headerRelationshipFiles) {
      const relationshipXml = await zip.file(relationshipFile)?.async('string');
      if (!relationshipXml) continue;
      const targetPattern = /\bTarget="media\/([^"]+)"/g;
      let targetMatch: RegExpExecArray | null;
      while ((targetMatch = targetPattern.exec(relationshipXml)) !== null) {
        const headerImagePath = `word/media/${targetMatch[1]}`;
        if (zip.file(headerImagePath)) zip.file(headerImagePath, scannedHeader);
      }
    }
  }

  journalImages.forEach((image) => zip.file(image.path, image.data));

  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

function parseBase64Image(base64String: string): Buffer | null {
  if (!base64String || typeof base64String !== 'string') return null;
  if (!base64String.startsWith('data:image/')) return null;

  try {
    const cleanBase64 = base64String.replace(/^data:image\/\w+;base64,/, "");
    return Buffer.from(cleanBase64, 'base64');
  } catch (err) {
    console.error("Failed to parse image", err);
    return null;
  }
}

export async function POST(request: Request) {
  try {
    const data = await request.json();

    if (!data || typeof data !== 'object') {
      return new Response(JSON.stringify({ error: 'Invalid payload' }), { status: 400 });
    }

    const isPreviewRequest = request.headers.get('x-report-preview') === 'true';

    if (hasSupabaseConfig && !isPreviewRequest) {
      const supabase = createClient(supabaseUrl!, supabaseKey!);
      const payload = {
        student_name: data.studentName || null,
        appendices: JSON.stringify(data.appendices) || null,
      };
      await supabase.from('reports').insert([payload]);
    }

    if (data.reportTemplate === 'scanned-template') {
      const filledTemplate = await fillReferenceTemplate(data);
      return new Response(filledTemplate, {
        status: 200,
        headers: {
          'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          'Content-Disposition': `attachment; filename="${OUTPUT_FILE_NAME}"`,
        },
      });
    }

    const reportFooterChildren = data.appendices?.reportFooter
      ? buildReportFooterBlock(data.appendices?.reportFooter)
      : [];
    const compactSectionLayout = data.sectionLayout === 'compact';
    const acknowledgementLineSpacing = 360;

    const documentSections: any[] = [
      {
        properties: {
          page: {
            size: { width: 12240, height: 20160 },
            margin: { top: 1440, bottom: 1440, left: 1440, right: 1440, header: 0, footer: 0 },
          },
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({ alignment: AlignmentType.CENTER, children: buildHeaderImage() }),
            ],
          }),
        },
        children: [
          ...buildCoverPage(data),
          new Paragraph({ pageBreakBefore: true }),
          ...buildAcknowledgementPage(data, compactSectionLayout, acknowledgementLineSpacing),
          new Paragraph({ pageBreakBefore: true }),
          ...buildTableOfContentsPage(data, compactSectionLayout, acknowledgementLineSpacing),
          new Paragraph({ pageBreakBefore: true }),
          ...buildSectionPage('INTRODUCTION', [
            { title: 'Background of the Organization', content: data.background },
            { title: 'Vision', content: data.vision },
            { title: 'Mission', content: data.mission },
            { title: 'Objectives', content: data.objectives, isBullet: true },
            { title: 'Core Values', content: data.coreValues, isBullet: true },
            { title: 'Products and Services Offered', content: data.services },
          ], compactSectionLayout, acknowledgementLineSpacing, data.organizationStructureImage),
          new Paragraph({ pageBreakBefore: !compactSectionLayout }),
          ...buildSectionPage('ORGANIZATION / COMPANY ANALYSIS', [
            { title: 'Strengths', content: data.strengths, isBullet: true },
            { title: 'Weaknesses', content: data.weaknesses, isBullet: true },
            { title: 'Opportunities', content: data.opportunities, isBullet: true },
            { title: 'Threats', content: data.threats, isBullet: true },
            { title: 'Recommendations for Improvement', content: data.recommendations, isBullet: true },
          ], compactSectionLayout, acknowledgementLineSpacing),
          new Paragraph({ pageBreakBefore: !compactSectionLayout }),
          ...buildSectionPage('TASKS AND DUTIES', [
            { title: 'Assigned Tasks and Responsibilities', content: data.tasks, isBullet: true },
            { title: 'Duties and Procedures Conformed', content: data.procedures, isBullet: true },
          ], compactSectionLayout, acknowledgementLineSpacing),
          new Paragraph({ pageBreakBefore: !compactSectionLayout }),
          ...buildSectionPage('CASE ANALYSIS', [
            { title: 'Issue / Problem 1', content: data.issue1 },
            { title: 'Strategy/Action Undertaken for Problem 1', content: data.issue1Action },
            { title: 'Issue / Problem 2', content: data.issue2 },
            { title: 'Strategy/Action Undertaken for Problem 2', content: data.issue2Action },
            { title: 'Lessons Learned from the Situations', content: data.lessons },
          ], compactSectionLayout, acknowledgementLineSpacing),
          new Paragraph({ pageBreakBefore: !compactSectionLayout }),
          ...buildSectionPage('REFLECTIONS', [
            { title: 'Self-Evaluation', content: data.selfEvaluation },
            { title: 'Relevancy of the Organization', content: data.relevancy },
          ], compactSectionLayout, acknowledgementLineSpacing),
          ...(data.schoolPlacement?.enabled
            ? [
              new Paragraph({ pageBreakBefore: true }),
              ...buildSchoolPlacementPages(data.schoolPlacement, compactSectionLayout, acknowledgementLineSpacing),
            ]
            : []),
        ],
      },
    ];

    const dailyJournalContent = buildDailyJournalAppendixPage(data.appendices);
    if (dailyJournalContent.length > 0) {
      documentSections.push({
        properties: {
          page: {
            size: { width: 12240, height: 20160 },
            margin: { top: 1440, bottom: 1440, left: 1440, right: 1440, header: 0, footer: 0 },
          },
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({ alignment: AlignmentType.CENTER, children: buildHeaderImage() }),
            ],
          }),
        },
        footers: reportFooterChildren.length > 0 ? {
          default: new Footer({
            children: reportFooterChildren,
          }),
        } : undefined,
        children: [
          ...(!compactSectionLayout ? [new Paragraph({ pageBreakBefore: true })] : []),
          ...dailyJournalContent,
        ],
      });
    }

    const remainingAppendixContent = buildOtherAppendixPages(data.appendices);
    if (remainingAppendixContent.length > 0) {
      documentSections.push({
        properties: {
          page: {
            size: { width: 12240, height: 20160 },
            margin: { top: 1440, bottom: 1440, left: 1440, right: 1440, header: 0, footer: 0 },
          },
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({ alignment: AlignmentType.CENTER, children: buildHeaderImage() }),
            ],
          }),
        },
        children: [
          ...(!compactSectionLayout ? [new Paragraph({ pageBreakBefore: true })] : []),
          ...remainingAppendixContent,
        ],
      });
    }

    if (data.reportTemplate === 'docx-template') {
      documentSections.splice(
        0,
        documentSections.length,
        ...buildDocxTemplateSections(data),
      );
    }

    if (data.reportTemplate === 'scanned-template') {
      documentSections.splice(
        0,
        documentSections.length,
        ...buildDocxTemplateSections(data, HD_TEMPLATE_HEADER, true),
      );
    }

    const doc = new Document({
      sections: documentSections,
    });

    const buffer = await Packer.toBuffer(doc);
    const docBuffer = Buffer.from(buffer);

    return new Response(docBuffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="${OUTPUT_FILE_NAME}"`,
      },
    });
  } catch (error) {
    console.error(error);
    return new Response(JSON.stringify({ error: 'Something went wrong while generating the report.' }), { status: 500 });
  }
}

function templateParagraph(text: string, options: {
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
  bold?: boolean;
  size?: number;
  before?: number;
  after?: number;
  pageBreakBefore?: boolean;
  justified?: boolean;
  bullet?: boolean;
} = {}): Paragraph[] {
  const paragraph = new Paragraph({
    alignment: options.alignment ?? (options.justified ? AlignmentType.JUSTIFIED : AlignmentType.LEFT),
    pageBreakBefore: options.pageBreakBefore,
    spacing: {
      before: options.before ?? 0,
      after: options.after ?? 0,
      line: 360,
      lineRule: 'auto',
    },
    indent: options.bullet ? { left: 720 } : undefined,
    bullet: options.bullet ? { level: 0 } : undefined,
    children: [new TextRun({
      text,
      bold: options.bold ?? false,
      size: options.size ?? 24,
      font: 'Times New Roman',
    })],
  });
  return [paragraph];
}

function buildTemplateCoverPage(data: any): Paragraph[] {
  const placementOrganizations = [
    data.trainingOrganization,
    data.schoolPlacement?.enabled ? data.schoolPlacement.trainingOrganization : '',
  ].filter(Boolean);
  const placementLocations = [
    data.trainingLocation,
    data.schoolPlacement?.enabled ? data.schoolPlacement.trainingLocation : '',
  ].filter(Boolean);

  const centered = (text: string, options: {
    before?: number;
    after?: number;
    bold?: boolean;
    size?: number;
    line?: number;
  } = {}) => new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: {
      before: options.before ?? 0,
      after: options.after ?? 360,
      line: options.line ?? 360,
      lineRule: 'auto',
    },
    children: [new TextRun({
      text,
      bold: options.bold ?? false,
      size: options.size ?? 32,
      font: 'Times New Roman',
    })],
  });

  return [
    centered('A Narrative Report on the', { before: 1440, after: 120, size: 32 }),
    centered('On-the-Job Training conducted at', { after: 240, size: 32 }),
    centered(placementOrganizations.join(' and '), { bold: true, after: 120, size: 32 }),
    centered(placementLocations.join(' and '), { after: 1440, size: 32 }),
    centered('Presented to the faculty of', { after: 120, size: 32 }),
    centered(data.collegeFaculty || '', { after: 1440, size: 32 }),
    centered('In partial fulfillment of the requirements for the degree of', { after: 120, size: 32 }),
    centered(data.degreeProgram || '', { after: 2160, size: 32 }),
    centered('Training Duration:', { after: 120, size: 32 }),
    centered(data.trainingDuration || '', { after: 1200, size: 32 }),
    centered('Submitted by:', { after: 120, size: 32 }),
    centered(data.studentName || '', { bold: true, after: 120, size: 32 }),
    centered(`${data.degreeProgram || ''}${data.studentYear ? ` - ${data.studentYear}` : ''}`, { after: 1440, size: 32 }),
    centered('Submitted to:', { after: 120, size: 32 }),
    centered(data.submittedToName || '', { bold: true, after: 120, size: 32 }),
    centered(data.submittedToTitle || '', { after: 0, size: 32 }),
  ];
}

function buildTemplateTableOfContents(data: any): any[] {
  const schoolEnabled = data.schoolPlacement?.enabled;
  const rows: [string, string][] = [
    ['Acknowledgement', 'i'],
    ['Table of Contents', 'ii'],
    ['List of Table/Figures', 'iii'],
    ['I. ' + String(data.trainingOrganization || 'TRAINING ORGANIZATION').toUpperCase(), ''],
    ['Introduction', '1'],
    ['  Background of the Organization', ''],
    ['  Vision', ''],
    ['  Mission', ''],
    ['  Objectives', ''],
    ['  Core Values', ''],
    ['  Products and Services Offered', ''],
    ['Organization / Company Analysis', '4'],
    ['Tasks and Duties', '6'],
    ['Case Analysis', '7'],
    ['Reflections', '8'],
  ];
  if (schoolEnabled) {
    rows.push(
      ['II. ' + String(data.schoolPlacement?.trainingOrganization || 'SECOND OJT PLACEMENT').toUpperCase(), ''],
      ['Introduction', '9'],
      ['  Background of the Organization', ''],
      ['  Vision', ''],
      ['  Mission', ''],
      ['  Objectives', ''],
      ['  Core Values', ''],
      ['  Products and Services Offered', ''],
      ['Organization / Company Analysis', '12'],
      ['Tasks and Duties', '14'],
      ['Case Analysis', '15'],
      ['Reflections', '16'],
      ['Appendices', '17'],
    );
  } else {
    rows.push(['Appendices', '9']);
  }

  return rows.map(([title, page]) => new Paragraph({
    spacing: { after: 60, line: 360, lineRule: 'auto' },
    children: [
      new TextRun({ text: `${title}${page ? ' ' + '-'.repeat(Math.max(8, 82 - title.length)) : ''}`, size: 24, font: 'Times New Roman', bold: /^(I\.|II\.|Organization\/Company Analysis|Tasks and Duties|Case Analysis|Reflections)/.test(title) }),
      ...(page ? [
        new TextRun({ text: `  ${page}`, size: 24, font: 'Times New Roman' }),
      ] : []),
    ],
  }));
}

function buildTemplateDailyJournal(appendices: AppendicesData | undefined): (Paragraph | Table)[] {
  if (!appendices?.dailyJournal?.length) return [];
  const children: (Paragraph | Table)[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 240, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: 'APPENDICES', bold: true, size: 24, font: 'Times New Roman' })],
    }),
  ];
  const layout = appendices.dailyJournalLayout === 'no-photo'
    ? 'no-photo'
    : appendices.dailyJournalLayout === 'report' ? 'report' : 'current';

  appendices.dailyJournal.forEach((week, index) => {
    if (layout === 'report') {
      const reportImages = week.reportLayoutImages || week.images || [];
      children.push(
        new Paragraph({
          pageBreakBefore: index > 0,
          spacing: { before: 120, after: 120, line: 360, lineRule: 'auto' },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: `WEEK ${week.weekNumber}`, bold: true, size: 24, font: 'Times New Roman' })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 60, after: 120, line: 360, lineRule: 'auto' },
          children: [new TextRun({ text: 'WEEKLY ACCOMPLISHMENT REPORT', bold: true, size: 24, font: 'Times New Roman' })],
        }),
        buildDailyReportTable(week.activities, reportImages),
      );
      return;
    }

    children.push(
      new Paragraph({
        pageBreakBefore: index > 0,
        spacing: { before: 120, after: 120, line: 360, lineRule: 'auto' },
        children: [new TextRun({ text: `Week ${week.weekNumber}`, bold: true, size: 24, font: 'Times New Roman' })],
      }),
      buildCurrentWeeklyTable(week.activities, week.totalHours),
    );

    if (layout === 'current') {
      children.push(
        new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 120, line: 360, lineRule: 'auto' },
        children: [new TextRun({ text: `Table ${week.weekNumber}: Journal Week ${week.weekNumber}`, bold: true, size: 24, font: 'Times New Roman' })],
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 120, line: 360, lineRule: 'auto' },
        children: [new TextRun({ text: week.narrative || 'No narrative provided.', size: 24, font: 'Times New Roman' })],
      }),
      );
    }

    if (layout === 'current') {
      const images = (week.currentLayoutImages || week.images || []).filter((image): image is string | AppendixImage => Boolean(image));
      if (images.length > 0) {
        children.push(buildImageGridTable(images));
        images.forEach((image, imageIndex) => {
          const detail = typeof image === 'string' ? '' : image.detail?.trim();
          children.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 60, after: 120, line: 360, lineRule: 'auto' },
            children: [new TextRun({
              text: `Figure ${imageIndex + 1}${detail ? `: ${detail}` : ''}`,
              size: 24,
              font: 'Times New Roman',
            })],
          }));
        });
      }
    }
  });

  return children;
}

function buildTemplateOrganizationPages(label: string, source: any, romanLabel = '', figureNumber = 1): Paragraph[] {
  const children: Paragraph[] = [
    new Paragraph({
      pageBreakBefore: true,
      alignment: AlignmentType.CENTER,
      spacing: { after: 120, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: `${romanLabel}${romanLabel ? '. ' : ''}${String(label || 'TRAINING ORGANIZATION').toUpperCase()}`, bold: true, size: 24, font: 'Times New Roman' })],
    }),
  ];

  const addPageHeading = (text: string) => {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 180, line: 360, lineRule: 'auto' },
        children: [new TextRun({ text, bold: true, size: 24, font: 'Times New Roman' })],
      }),
    );
  };
  const addSubheading = (text: string) => {
    children.push(
      new Paragraph({
        spacing: { before: 60, after: 60, line: 360, lineRule: 'auto' },
        children: [new TextRun({ text: `${text}:`, bold: true, size: 24, font: 'Times New Roman' })],
      }),
    );
  };
  const addText = (text: string, bullet = false) => {
    const lines = ensureArray(text);
    if (!lines.length) return;
    lines.forEach((line) => children.push(
      new Paragraph({
        alignment: bullet ? AlignmentType.LEFT : AlignmentType.JUSTIFIED,
        indent: bullet ? { left: 720 } : undefined,
        bullet: bullet ? { level: 0 } : undefined,
        spacing: { before: 0, after: 0, line: 360, lineRule: 'auto' },
        children: [new TextRun({ text: line, size: 24, font: 'Times New Roman' })],
      }),
    ));
  };

  addPageHeading('INTRODUCTION');
  addSubheading('Background of the Organization');
  addText(source.background);
  addSubheading('Vision');
  addText(source.vision);
  addSubheading('Mission');
  addText(source.mission);
  addSubheading('Objectives');
  addText(source.objectives, true);
  addSubheading('Core Values');
  addText(source.coreValues, true);
  addSubheading('Products and Services Offered');
  addText(source.services, true);

  const imageBuffer = parseBase64Image(source.organizationStructureImage || '');
  if (imageBuffer) {
    children.push(
      ...templateParagraph('Organization Structure', { alignment: AlignmentType.CENTER, bold: true, before: 120 }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new ImageRun({ type: 'png', data: imageBuffer, transformation: { width: 520, height: 360 } })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 120, line: 360, lineRule: 'auto' },
        children: [new TextRun({ text: `Figure ${figureNumber}: Organizational Chart`, size: 24, font: 'Times New Roman' })],
      }),
    );
  }

  children.push(
    new Paragraph({ pageBreakBefore: true }),
    ...templateParagraph('ORGANIZATION / COMPANY ANALYSIS', { alignment: AlignmentType.CENTER, bold: true }),
  );
  addSubheading('Strengths');
  addText(source.strengths, true);
  addSubheading('Weaknesses');
  addText(source.weaknesses, true);
  addSubheading('Opportunities');
  addText(source.opportunities, true);
  addSubheading('Threats');
  addText(source.threats, true);
  addSubheading('Recommendations for Improvement');
  addText(source.recommendations, true);

  children.push(
    new Paragraph({ pageBreakBefore: true }),
    ...templateParagraph('TASKS AND DUTIES', { alignment: AlignmentType.CENTER, bold: true }),
  );
  addSubheading('Assigned Tasks and Responsibilities');
  addText(source.tasks, true);
  addSubheading('Duties and Procedures Conformed');
  addText(source.procedures, true);

  children.push(
    new Paragraph({ pageBreakBefore: true }),
    ...templateParagraph('CASE ANALYSIS', { alignment: AlignmentType.CENTER, bold: true }),
  );
  addSubheading('Issue / Problem 1');
  addText(source.issue1);
  addSubheading('Strategy/Action Undertaken for Problem 1');
  addText(source.issue1Action);
  addSubheading('Issue / Problem 2');
  addText(source.issue2);
  addSubheading('Strategy/Action Undertaken for Problem 2');
  addText(source.issue2Action);
  addSubheading('Lessons Learned from the Situations');
  addText(source.lessons);

  children.push(
    new Paragraph({ pageBreakBefore: true }),
    ...templateParagraph('REFLECTIONS', { alignment: AlignmentType.CENTER, bold: true }),
  );
  addSubheading('Self-Evaluation');
  addText(source.selfEvaluation);
  addSubheading('Relevancy of the Organization');
  addText(source.relevancy);

  return children;
}

function buildDocxTemplateSections(
  data: any,
  headerImage: Buffer | null = HEADER_IMAGE,
  scannedTemplate = false,
): any[] {
  const children: (Paragraph | Table)[] = [
    ...buildTemplateCoverPage(data),
    new Paragraph({ pageBreakBefore: true }),
    ...buildAcknowledgementPage(data, false, 360),
    new Paragraph({ pageBreakBefore: true }),
    ...templateParagraph('TABLE OF CONTENTS', { alignment: AlignmentType.CENTER, bold: true, after: 180 }),
    ...buildTemplateTableOfContents(data),
    new Paragraph({ pageBreakBefore: true }),
    ...templateParagraph('LIST OF TABLES / FIGURES', { alignment: AlignmentType.CENTER, bold: true, after: 180 }),
    ...templateParagraph('LIST OF TABLES', { bold: true, after: 60 }),
  ];

  (data.appendices?.dailyJournal || []).forEach((week: DailyJournalWeek, index: number) => {
    children.push(...templateParagraph(`Table ${index + 1}: Journal Week ${week.weekNumber}`, {}));
  });

  children.push(
    ...templateParagraph('LIST OF FIGURES', { bold: true, before: 120, after: 60 }),
  );

  if (data.organizationStructureImage) {
    children.push(...templateParagraph('Figure 1: Organizational Chart', {}));
  }
  if (data.schoolPlacement?.enabled && data.schoolPlacement.organizationStructureImage) {
    children.push(...templateParagraph('Figure 2: School Placement Organizational Chart', {}));
  }
  const hasJournalPhotos = (data.appendices?.dailyJournal || []).some((week: DailyJournalWeek) =>
    (week.currentLayoutImages || week.images || []).some((image) => Boolean(image)),
  );
  if (hasJournalPhotos) {
    children.push(...templateParagraph('Figure 3: Weekly documentation photos', {}));
  }
  if (data.appendices?.primeReportImages?.length) {
    children.push(...templateParagraph('Figure 4: PRIME report photos', {}));
  }

  children.push(...buildTemplateOrganizationPages(data.trainingOrganization, data, 'I', 1));
  if (data.schoolPlacement?.enabled) {
    children.push(...buildTemplateOrganizationPages(data.schoolPlacement.trainingOrganization || 'SCHOOL OJT / SECOND PLACEMENT', data.schoolPlacement, 'II', 2));
  }

  const journalChildren = buildTemplateDailyJournal(data.appendices);
  if (journalChildren.length > 0) {
    children.push(new Paragraph({ pageBreakBefore: true }), ...journalChildren);
  }
  const otherAppendices = buildOtherAppendixPages(data.appendices);
  if (otherAppendices.length > 0) {
    children.push(...otherAppendices);
  }

  return [{
    properties: {
      page: {
        size: scannedTemplate
          ? { width: 12240, height: 20160 }
          : { width: 12242, height: 20163 },
        margin: scannedTemplate
          ? { top: 762000, bottom: 914400, left: 914400, right: 914400, header: 449580, footer: 449580 }
          : { top: 1440, bottom: 1440, left: 1440, right: 1440, header: 709, footer: 709 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: buildHeaderImage(
            headerImage,
            scannedTemplate ? 624 : 624,
            scannedTemplate ? 109 : 125,
          ),
        })],
      }),
    },
    children,
  }];
}

function buildCoverPage(data: any): Paragraph[] {
  const org = [data.trainingOrganization, data.schoolPlacement?.enabled ? data.schoolPlacement.trainingOrganization : '']
    .filter(Boolean)
    .join(' and ');
  const location = [data.trainingLocation, data.schoolPlacement?.enabled ? data.schoolPlacement.trainingLocation : '']
    .filter(Boolean)
    .join(' and ');
  const faculty = data.collegeFaculty || '';
  const degree = data.degreeProgram || '';
  const student = data.studentName || '';
  const studentDegree = data.degreeProgram || '';
  const adviserName = data.submittedToName || '';
  const adviserTitle = data.submittedToTitle || '';
  const FRONT_PAGE_TEXT_SIZE = 20;

  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 1440, after: 360 },
      children: [new TextRun({ text: 'A Narrative Report on the', size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 360, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: 'On-the-Job Training conducted at', size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 180, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: org, size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 1440, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: location, size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 360, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: 'Presented to the faculty of', size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 1440, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: faculty, size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 360, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: 'In partial fulfillment of the requirements for the degree of', size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 2160, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: degree, size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 180, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: 'Submitted by:', size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 180, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: student, size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 1080, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: studentDegree, size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 180, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: 'Submitted to:', size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 180, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: adviserName, size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 0, line: 360, lineRule: 'auto' },
      children: [new TextRun({ text: adviserTitle, size: FRONT_PAGE_TEXT_SIZE, font: 'Times New Roman', bold: false })],
    }),
  ];
}

function buildAcknowledgementPage(data: any, compact = false, lineSpacing = 240): Paragraph[] {
  const student = data.studentName || '';
  const degree = data.degreeProgram || '';
  const university = 'Jose Rizal Memorial State University';
  const paragraphs: Paragraph[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
      children: [new TextRun({ text: 'ACKNOWLEDGEMENT', bold: true, size: 24, font: 'Times New Roman' })],
    }),
  ];

  const rawContent = data.acknowledgement || '';
  const textParagraphs = ensureArray(rawContent);

  if (textParagraphs.length > 0) {
    textParagraphs.forEach((pText) => {
      paragraphs.push(
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          indent: { firstLine: 720 },
          spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
          children: [new TextRun({ text: pText, size: 24, font: 'Times New Roman' })],
        })
      );
    });
  } else {
    paragraphs.push(
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        indent: { firstLine: 720 },
        spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
        children: [new TextRun({ text: 'No acknowledgement provided.', size: 24, font: 'Times New Roman' })],
      })
    );
  }

  // Signature Block matching the uploaded image format exactly
  paragraphs.push(
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
      children: [new TextRun({ text: student, bold: true, size: 24, font: 'Times New Roman' })],
    }),
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
      children: [new TextRun({ text: degree, size: 24, font: 'Times New Roman' })],
    }),
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
      children: [new TextRun({ text: university, size: 24, font: 'Times New Roman' })],
    })
  );

  return paragraphs;
}

function normalizeTOCKey(value: string): string {
  return value
    .replace(/^\s*\d+(?:\.\d+)*[.)]?\s*/, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\b(and|of|the|for|from|to|a)\b/g, ' ')
    .trim();
}

const ESTIMATED_PAGE_LINES = 44;
const BODY_CHARACTERS_PER_LINE = 84;

function estimateWrappedLines(text: string, charactersPerLine = BODY_CHARACTERS_PER_LINE): number {
  return text.split(/\r?\n/).reduce((total, line) => {
    return total + Math.max(1, Math.ceil(line.length / charactersPerLine));
  }, 0);
}

function estimateSectionLineCount(sectionData: SectionData[], hasImage = false): number {
  let totalLines = 1;

  sectionData.forEach((section) => {
    const text = typeof section.content === 'string'
      ? section.content
      : Array.isArray(section.content)
        ? section.content.join('\n')
        : String(section.content ?? '');
    const lines = ensureArray(text);

    totalLines += estimateWrappedLines(`${section.title}:`);
    if (section.isBullet) {
      totalLines += lines.reduce((count, line) => count + estimateWrappedLines(line), 0);
    } else if (lines.length > 0) {
      totalLines += estimateWrappedLines(lines.join(' '));
    }
  });

  if (hasImage) totalLines += 17;
  return totalLines;
}

function estimateSectionPageCount(sectionData: SectionData[], hasImage = false): number {
  return Math.max(1, Math.ceil(estimateSectionLineCount(sectionData, hasImage) / ESTIMATED_PAGE_LINES));
}

function addSubsectionPageMap(
  pageMap: Record<string, number>,
  sectionNumber: string,
  sectionData: SectionData[],
  startPage: number,
  compact: boolean,
) {
  let lineCursor = 1;

  sectionData.forEach((section, index) => {
    pageMap[`${sectionNumber}.${index + 1}`] = compact
      ? startPage + Math.floor(lineCursor / ESTIMATED_PAGE_LINES)
      : startPage;
    lineCursor += estimateSectionLineCount([section]);
  });
}

function estimateAcknowledgementLineCount(data: any): number {
  const paragraphs = ensureArray(data.acknowledgement || '');
  return 4 + paragraphs.reduce((total, paragraph) => total + estimateWrappedLines(paragraph), 0);
}

function estimateTOCLineCount(data: any): number {
  const entries = ensureArray(data.tableOfContents || '');
  const baseEntryCount = entries.filter((entry) => !isSchoolTOCLine(entry)).length || 26;
  return 1 + baseEntryCount + (data.schoolPlacement?.enabled ? 25 : 0);
}

function isSchoolTOCLine(title: string): boolean {
  return /school ojt/i.test(title)
    || /^(?:6|7|8|9|10)\.\d+(?:\.\d+)?\s+/.test(title)
    || /^(6\.\s*introduction|7\.\s*organization|8\.\s*tasks|9\.\s*case|10\.\s*reflections)/i.test(title);
}

function estimateAppendixPageCount(appendicesData?: AppendicesData): number {
  if (!appendicesData) return 1;

  let pages = 1;

  if (appendicesData.dailyJournal && appendicesData.dailyJournal.length > 0) {
    pages += appendicesData.dailyJournal.length * 2;
  }

  const appendixKeys = [
    'certParticipation',
    'primeNarrative',
    'resume',
    'grades',
    'medicalWaiver',
    'letterAcceptance',
    'dtr',
    'ratingSheet',
    'certCompletion',
  ] as const;

  appendixKeys.forEach((key) => {
    if (appendicesData[key]) pages += 1;
  });

  return Math.max(1, pages);
}

function getSectionPageMap(data: any): Record<string, number> {
  const compact = data.sectionLayout === 'compact';
  const acknowledgementLines = estimateAcknowledgementLineCount(data);
  const tocLines = estimateTOCLineCount(data);
  const acknowledgementPages = Math.max(1, Math.ceil(acknowledgementLines / ESTIMATED_PAGE_LINES));
  const tocPages = Math.max(1, Math.ceil(tocLines / ESTIMATED_PAGE_LINES));
  const narrativeStartPage = 2 + acknowledgementPages + tocPages;
  let pageCursor = compact
    ? narrativeStartPage
    : 2
      + acknowledgementPages
      + tocPages;
  let lineCursor = 0;
  const pageMap: Record<string, number> = {};
  const sectionDefinitions = [
    { key: 'introduction', value: [
      { title: 'Background of the Organization', content: data.background },
      { title: 'Vision', content: data.vision },
      { title: 'Mission', content: data.mission },
      { title: 'Objectives', content: data.objectives, isBullet: true },
      { title: 'Core Values', content: data.coreValues, isBullet: true },
      { title: 'Products and Services Offered', content: data.services },
    ], image: Boolean(data.organizationStructureImage) },
    { key: 'organization company analysis', value: [
      { title: 'Strengths', content: data.strengths, isBullet: true },
      { title: 'Weaknesses', content: data.weaknesses, isBullet: true },
      { title: 'Opportunities', content: data.opportunities, isBullet: true },
      { title: 'Threats', content: data.threats, isBullet: true },
      { title: 'Recommendations for Improvement', content: data.recommendations, isBullet: true },
    ], image: false },
    { key: 'tasks and duties', value: [
      { title: 'Assigned Tasks and Responsibilities', content: data.tasks, isBullet: true },
      { title: 'Duties and Procedures Conformed', content: data.procedures, isBullet: true },
    ], image: false },
    { key: 'case analysis', value: [
      { title: 'Issue / Problem 1', content: data.issue1 },
      { title: 'Strategy/Action Undertaken for Problem 1', content: data.issue1Action },
      { title: 'Issue / Problem 2', content: data.issue2 },
      { title: 'Strategy/Action Undertaken for Problem 2', content: data.issue2Action },
      { title: 'Lessons Learned from the Situations', content: data.lessons },
    ], image: false },
    { key: 'reflections', value: [
      { title: 'Self-Evaluation', content: data.selfEvaluation },
      { title: 'Relevancy of the Organization', content: data.relevancy },
    ], image: false },
  ];

  sectionDefinitions.forEach((section, index) => {
    const lineCount = estimateSectionLineCount(section.value, section.image);
    pageMap[section.key] = compact
      ? narrativeStartPage + Math.floor(lineCursor / ESTIMATED_PAGE_LINES)
      : pageCursor;
    addSubsectionPageMap(pageMap, String(index + 1), section.value, pageMap[section.key], compact);

    if (compact) {
      lineCursor += lineCount;
    } else {
      pageCursor += estimateSectionPageCount(section.value, section.image);
    }
  });

  if (data.schoolPlacement?.enabled) {
    const schoolDefinitions = [
      { key: 'school ojt introduction', value: [
        { title: 'School/OJT Unit', content: data.schoolPlacement.trainingOrganization },
        { title: 'Location', content: data.schoolPlacement.trainingLocation },
        { title: 'Background of the Organization', content: data.schoolPlacement.background },
        { title: 'Vision', content: data.schoolPlacement.vision },
        { title: 'Mission', content: data.schoolPlacement.mission },
        { title: 'Objectives', content: data.schoolPlacement.objectives, isBullet: true },
        { title: 'Core Values', content: data.schoolPlacement.coreValues, isBullet: true },
        { title: 'Products and Services Offered', content: data.schoolPlacement.services },
      ], image: Boolean(data.schoolPlacement.organizationStructureImage) },
      { key: 'school ojt organization analysis', value: [
        { title: 'Strengths', content: data.schoolPlacement.strengths, isBullet: true },
        { title: 'Weaknesses', content: data.schoolPlacement.weaknesses, isBullet: true },
        { title: 'Opportunities', content: data.schoolPlacement.opportunities, isBullet: true },
        { title: 'Threats', content: data.schoolPlacement.threats, isBullet: true },
        { title: 'Recommendations for Improvement', content: data.schoolPlacement.recommendations, isBullet: true },
      ], image: false },
      { key: 'school ojt tasks and duties', value: [
        { title: 'Assigned Tasks and Responsibilities', content: data.schoolPlacement.tasks, isBullet: true },
        { title: 'Duties and Procedures Conformed', content: data.schoolPlacement.procedures, isBullet: true },
      ], image: false },
      { key: 'school ojt case analysis', value: [
        { title: 'Issue / Problem 1', content: data.schoolPlacement.issue1 },
        { title: 'Strategy/Action Undertaken for Problem 1', content: data.schoolPlacement.issue1Action },
        { title: 'Issue / Problem 2', content: data.schoolPlacement.issue2 },
        { title: 'Strategy/Action Undertaken for Problem 2', content: data.schoolPlacement.issue2Action },
        { title: 'Lessons Learned from the Situations', content: data.schoolPlacement.lessons },
      ], image: false },
      { key: 'school ojt reflections', value: [
        { title: 'Self-Evaluation', content: data.schoolPlacement.selfEvaluation },
        { title: 'Relevancy of the Organization', content: data.schoolPlacement.relevancy },
      ], image: false },
    ];

    schoolDefinitions.forEach((section, index) => {
      const estimatedValue = index === 0
        ? [{ title: 'School OJT Narrative (Second Placement)', content: '' }, ...section.value]
        : section.value;
      const lineCount = estimateSectionLineCount(estimatedValue, section.image);
      const schoolSectionNumber = String(index + 6);
      pageMap[section.key] = compact
        ? narrativeStartPage + Math.floor(lineCursor / ESTIMATED_PAGE_LINES)
        : pageCursor;
      addSubsectionPageMap(pageMap, schoolSectionNumber, section.value, pageMap[section.key], compact);

      if (compact) {
        lineCursor += lineCount;
      } else {
        pageCursor += estimateSectionPageCount(estimatedValue, section.image);
      }
    });

    pageMap['school ojt narrative'] = pageMap['school ojt introduction'];
  }

  pageMap.appendices = compact
    ? narrativeStartPage + Math.ceil(lineCursor / ESTIMATED_PAGE_LINES)
    : pageCursor;
  return pageMap;
}

function buildTableOfContentsPage(data: any, compact = false, lineSpacing = 240): Paragraph[] {
  const rawContent = data.tableOfContents || '';
  const sectionPageMap = getSectionPageMap(data);

  const fallbackEntries = [
    { title: '1. Introduction', page: String(sectionPageMap.introduction || 4) },
    { title: '1.1 Background of the Organization', page: String(sectionPageMap.introduction || 4) },
    { title: '1.2 Vision', page: String(sectionPageMap.introduction || 4) },
    { title: '1.3 Mission', page: String(sectionPageMap.introduction || 4) },
    { title: '1.4 Objectives', page: String(sectionPageMap.introduction || 4) },
    { title: '1.5 Core Values', page: String(sectionPageMap.introduction || 4) },
    { title: '1.6 Products and Services Offered', page: String(sectionPageMap.introduction || 4) },
    { title: '2. Organization / Company Analysis', page: String(sectionPageMap['organization company analysis'] || 5) },
    { title: '2.1 Strengths', page: String(sectionPageMap['organization company analysis'] || 5) },
    { title: '2.2 Weaknesses', page: String(sectionPageMap['organization company analysis'] || 5) },
    { title: '2.3 Opportunities', page: String(sectionPageMap['organization company analysis'] || 5) },
    { title: '2.4 Threats', page: String(sectionPageMap['organization company analysis'] || 5) },
    { title: '2.5 Recommendations for Improvement', page: String(sectionPageMap['organization company analysis'] || 5) },
    { title: '3. Tasks and Duties', page: String(sectionPageMap['tasks and duties'] || 6) },
    { title: '3.1 Assigned Tasks and Responsibilities', page: String(sectionPageMap['tasks and duties'] || 6) },
    { title: '3.2 Duties and Procedures Conformed', page: String(sectionPageMap['tasks and duties'] || 6) },
    { title: '4. Case Analysis', page: String(sectionPageMap['case analysis'] || 7) },
    { title: '4.1 Issue / Problem 1', page: String(sectionPageMap['case analysis'] || 7) },
    { title: '4.2 Strategy/Action Undertaken for Problem 1', page: String(sectionPageMap['case analysis'] || 7) },
    { title: '4.3 Issue / Problem 2', page: String(sectionPageMap['case analysis'] || 7) },
    { title: '4.4 Strategy/Action Undertaken for Problem 2', page: String(sectionPageMap['case analysis'] || 7) },
    { title: '4.5 Lessons Learned from the Situations', page: String(sectionPageMap['case analysis'] || 7) },
    { title: '5. Reflections', page: String(sectionPageMap.reflections || 8) },
    { title: '5.1 Self-Evaluation', page: String(sectionPageMap.reflections || 8) },
    { title: '5.2 Relevancy of the Organization', page: String(sectionPageMap.reflections || 8) },
    { title: '6. Appendices', page: String(sectionPageMap.appendices || 9) },
  ];

  const parsedEntries = ensureArray(rawContent).reduce<{ title: string; page: string }[]>((list, line) => {
    const trimmed = line.trim();
    if (!trimmed) return list;

    const match = trimmed.match(/^(.*?)(?:\s+)(\d+)$/) || trimmed.match(/^(.*)$/);
    const extractedTitle = (match?.[1] || trimmed)
      .replace(/\s*[:.-]+\s*$/, '')
      .trim();
    const extractedPage = match?.[2] || '';

    if (extractedTitle) {
      list.push({
        title: extractedTitle,
        page: extractedPage,
      });
    }

    return list;
  }, []);

  let entries = parsedEntries.length > 0 ? parsedEntries : fallbackEntries;
  if (data.schoolPlacement?.enabled) {
    entries = entries.filter((entry) => !isSchoolTOCLine(entry.title));
    const schoolTOCEntries = [
      { title: '6. Introduction', page: String(sectionPageMap['school ojt introduction'] || 9) },
      { title: '6.1 Background of the Organization', page: String(sectionPageMap['school ojt introduction'] || 9) },
      { title: '6.2 Vision', page: String(sectionPageMap['school ojt introduction'] || 9) },
      { title: '6.3 Mission', page: String(sectionPageMap['school ojt introduction'] || 9) },
      { title: '6.4 Objectives', page: String(sectionPageMap['school ojt introduction'] || 9) },
      { title: '6.5 Core Values', page: String(sectionPageMap['school ojt introduction'] || 9) },
      { title: '6.6 Products and Services Offered', page: String(sectionPageMap['school ojt introduction'] || 9) },
      { title: '7. Organization / Company Analysis', page: String(sectionPageMap['school ojt organization analysis'] || 9) },
      { title: '7.1 Strengths', page: String(sectionPageMap['school ojt organization analysis'] || 9) },
      { title: '7.2 Weaknesses', page: String(sectionPageMap['school ojt organization analysis'] || 9) },
      { title: '7.3 Opportunities', page: String(sectionPageMap['school ojt organization analysis'] || 9) },
      { title: '7.4 Threats', page: String(sectionPageMap['school ojt organization analysis'] || 9) },
      { title: '7.5 Recommendations for Improvement', page: String(sectionPageMap['school ojt organization analysis'] || 9) },
      { title: '8. Tasks and Duties', page: String(sectionPageMap['school ojt tasks and duties'] || 9) },
      { title: '8.1 Assigned Tasks and Responsibilities', page: String(sectionPageMap['school ojt tasks and duties'] || 9) },
      { title: '8.2 Duties and Procedures Conformed', page: String(sectionPageMap['school ojt tasks and duties'] || 9) },
      { title: '9. Case Analysis', page: String(sectionPageMap['school ojt case analysis'] || 9) },
      { title: '9.1 Issue / Problem 1', page: String(sectionPageMap['school ojt case analysis'] || 9) },
      { title: '9.2 Strategy/Action Undertaken for Problem 1', page: String(sectionPageMap['school ojt case analysis'] || 9) },
      { title: '9.3 Issue / Problem 2', page: String(sectionPageMap['school ojt case analysis'] || 9) },
      { title: '9.4 Strategy/Action Undertaken for Problem 2', page: String(sectionPageMap['school ojt case analysis'] || 9) },
      { title: '9.5 Lessons Learned from the Situations', page: String(sectionPageMap['school ojt case analysis'] || 9) },
      { title: '10. Reflections', page: String(sectionPageMap['school ojt reflections'] || 9) },
      { title: '10.1 Self-Evaluation', page: String(sectionPageMap['school ojt reflections'] || 9) },
      { title: '10.2 Relevancy of the Organization', page: String(sectionPageMap['school ojt reflections'] || 9) },
    ];
    const appendicesIndex = entries.findIndex((entry) => normalizeTOCKey(entry.title) === 'appendices');
    if (appendicesIndex >= 0) entries[appendicesIndex] = { ...entries[appendicesIndex], title: '11. Appendices' };
    entries.splice(appendicesIndex >= 0 ? appendicesIndex : entries.length, 0, ...schoolTOCEntries);
  } else {
    entries = entries.filter((entry) => !isSchoolTOCLine(entry.title));
  }

  const appendixEntries = [
    ['Appendix A - Daily Journal', 12],
    ['Appendix B - Certificate of Participation', 13],
    ['Appendix C - PRIME Narrative Report', 14],
    ['Appendix D - Resume', 15],
    ['Appendix E - Grades', 16],
    ['Appendix F - Medical Waiver', 17],
    ['Appendix G - Acceptance Letter', 18],
    ['Appendix H - DTR', 19],
    ['Appendix I - Rating Sheet', 20],
    ['Appendix J - Certificate of Completion', 21],
  ] as const;
  const appendicesHeadingIndex = entries.findIndex((entry) => normalizeTOCKey(entry.title) === 'appendices');
  const formattedAppendixEntries = appendixEntries.map(([title, page]) => ({
    title,
    page: String(page),
  }));
  if (appendicesHeadingIndex >= 0) {
    entries.splice(appendicesHeadingIndex + 1, 0, ...formattedAppendixEntries);
  } else {
    entries.push(...formattedAppendixEntries);
  }

  const paragraphs: Paragraph[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 240, line: lineSpacing, lineRule: 'auto' },
      children: [new TextRun({ text: 'Table of Contents', bold: true, size: 28, font: 'Times New Roman' })],
    }),
  ];

  entries.forEach((entry) => {
    const titleText = entry.title;
    const normalizedTitle = normalizeTOCKey(titleText);
    const sectionNumber = titleText.match(/^\s*(\d+)(?:\.\d+)*[.)]?\s+/)?.[1];
    const numberedSectionKeys: Record<string, string> = {
      '1': 'introduction',
      '2': 'organization company analysis',
      '3': 'tasks and duties',
      '4': 'case analysis',
      '5': 'reflections',
      '6': 'appendices',
    };
    const numberedPage = sectionNumber ? sectionPageMap[numberedSectionKeys[sectionNumber]] : undefined;
    const subsectionMatch = titleText.match(/^\s*(\d+)\.(\d+)/);
    const subsectionPage = subsectionMatch
      ? sectionPageMap[`${subsectionMatch[1]}.${subsectionMatch[2]}`]
      : undefined;
    const schoolMainNumber = titleText.match(/^\s*(\d+)(?:\.\d+)*[.)]?\s+/)?.[1];
    const schoolSectionKeys: Record<string, string> = {
      '6': 'school ojt introduction',
      '7': 'school ojt organization analysis',
      '8': 'school ojt tasks and duties',
      '9': 'school ojt case analysis',
      '10': 'school ojt reflections',
    };
    const schoolPlacementPage = data.schoolPlacement?.enabled && !normalizedTitle.includes('appendices')
      ? sectionPageMap[schoolSectionKeys[schoolMainNumber || '']]
      : undefined;
    const matchedPage = Object.entries(sectionPageMap).find(([key]) => {
      const normalizedKey = normalizeTOCKey(key);
      return normalizedKey === normalizedTitle
        || normalizedTitle.startsWith(normalizedKey)
        || normalizedKey.startsWith(normalizedTitle);
    })?.[1];
    const mappedPage = subsectionPage ?? schoolPlacementPage ?? matchedPage ?? numberedPage;
    const resolvedPage = mappedPage ? String(mappedPage) : entry.page;
    const isSubsection = /^\d+\.\d+\s+/.test(titleText);
    const showPageNumber = (compact || !isSubsection) && Boolean(resolvedPage);
    const fillLength = Math.max(4, showPageNumber ? 72 - titleText.length : 88 - titleText.length);
    const filler = '.'.repeat(fillLength);
    const indent = isSubsection ? 500 : 180;

    const pageText = showPageNumber ? [
      new TextRun({ text: '\t', size: 22, font: 'Times New Roman' }),
      new TextRun({ text: resolvedPage, size: 22, font: 'Times New Roman', bold: false }),
    ] : [];

    paragraphs.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        indent: { left: indent },
        tabStops: [{ type: 'right', position: 9000 }],
        spacing: { before: 0, after: 4, line: lineSpacing, lineRule: 'auto' },
        children: [
          new TextRun({ text: `${titleText}${filler}`, size: 22, font: 'Times New Roman' }),
          ...pageText,
        ],
      })
    );
  });

  return paragraphs;
}

function buildSectionPage(title: string, sections: SectionData[], compact = false, lineSpacing = 240, imageData?: string): Paragraph[] {
  const children: any[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
      children: [new TextRun({ text: title, bold: true, size: 24, font: 'Times New Roman' })],
    }),
  ];

  sections.forEach((section) => {
    const text = typeof section.content === 'string' ? section.content : Array.isArray(section.content) ? section.content.join('\n') : String(section.content ?? '');
    const contentLines = ensureArray(text);

    if (section.title) {
      children.push(
        new Paragraph({
          spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
          children: [new TextRun({ text: `${section.title}:`, bold: true, size: 24, font: 'Times New Roman' })],
        })
      );
    }

    if (section.isBullet) {
      contentLines.forEach((line) => {
        children.push(
          new Paragraph({
            bullet: { level: 0 },
            indent: { left: 720 },
            spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
            children: [new TextRun({ text: line, size: 24, font: 'Times New Roman' })],
          })
        );
      });
    } else if (contentLines.length > 0) {
      const combinedText = contentLines.join(' ');
      children.push(
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
          children: [new TextRun({ text: combinedText, size: 24, font: 'Times New Roman' })],
        })
      );
    }
  });

  const imageBuffer = parseBase64Image(imageData || '');
  if (imageBuffer) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 60 },
        children: [new TextRun({ text: 'Organization Structure', bold: true, size: 24, font: 'Times New Roman' })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 120 },
        children: [new ImageRun({ type: 'png', data: imageBuffer, transformation: { width: 520, height: 360 } })],
      })
    );
  }

  return children;
}

function buildSchoolPlacementPages(placement: any, compact = false, lineSpacing = 240): Paragraph[] {
  const pages: Paragraph[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0, line: lineSpacing, lineRule: 'auto' },
      children: [new TextRun({ text: 'SCHOOL OJT NARRATIVE (SECOND PLACEMENT)', bold: true, size: 24, font: 'Times New Roman' })],
    }),
  ];
  const sections = [
    {
      title: 'SCHOOL OJT - INTRODUCTION',
      content: [
        { title: 'School/OJT Unit', content: placement.trainingOrganization },
        { title: 'Location', content: placement.trainingLocation },
        { title: 'Background of the Organization', content: placement.background },
        { title: 'Vision', content: placement.vision },
        { title: 'Mission', content: placement.mission },
        { title: 'Objectives', content: placement.objectives, isBullet: true },
        { title: 'Core Values', content: placement.coreValues, isBullet: true },
        { title: 'Products and Services Offered', content: placement.services },
      ],
      image: placement.organizationStructureImage,
    },
    {
      title: 'SCHOOL OJT - ORGANIZATION ANALYSIS',
      content: [
        { title: 'Strengths', content: placement.strengths, isBullet: true },
        { title: 'Weaknesses', content: placement.weaknesses, isBullet: true },
        { title: 'Opportunities', content: placement.opportunities, isBullet: true },
        { title: 'Threats', content: placement.threats, isBullet: true },
        { title: 'Recommendations for Improvement', content: placement.recommendations, isBullet: true },
      ],
    },
    {
      title: 'SCHOOL OJT - TASKS AND DUTIES',
      content: [
        { title: 'Assigned Tasks and Responsibilities', content: placement.tasks, isBullet: true },
        { title: 'Duties and Procedures Conformed', content: placement.procedures, isBullet: true },
      ],
    },
    {
      title: 'SCHOOL OJT - CASE ANALYSIS',
      content: [
        { title: 'Issue / Problem 1', content: placement.issue1 },
        { title: 'Strategy/Action Undertaken for Problem 1', content: placement.issue1Action },
        { title: 'Issue / Problem 2', content: placement.issue2 },
        { title: 'Strategy/Action Undertaken for Problem 2', content: placement.issue2Action },
        { title: 'Lessons Learned from the Situations', content: placement.lessons },
      ],
    },
    {
      title: 'SCHOOL OJT - REFLECTIONS',
      content: [
        { title: 'Self-Evaluation', content: placement.selfEvaluation },
        { title: 'Relevancy of the Organization', content: placement.relevancy },
      ],
    },
  ];

  sections.forEach((section, index) => {
    if (index > 0 && !compact) pages.push(new Paragraph({ pageBreakBefore: true }));
    pages.push(...buildSectionPage(section.title, section.content, compact, lineSpacing, section.image));
  });

  return pages;
}

function buildImageGridTable(images: (string | AppendixImage)[]) {
  const rows: TableRow[] = [];
  const invisibleBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const tableBorders = { top: invisibleBorder, bottom: invisibleBorder, left: invisibleBorder, right: invisibleBorder, insideHorizontal: invisibleBorder, insideVertical: invisibleBorder };

  for (let i = 0; i < images.length; i += 2) {
    const img1 = images[i];
    const img2 = images[i + 1];

    const cell1Children: any[] = [];
    const cell2Children: any[] = [];

    if (img1) {
      const isObj1 = typeof img1 === 'object' && img1 !== null;
      const base64_1 = isObj1 ? (img1 as AppendixImage).base64 : (img1 as string);
      const detail_1 = isObj1 ? (img1 as AppendixImage).detail : '';
      const buf1 = parseBase64Image(base64_1);

      if (buf1) {
        cell1Children.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 60 },
            children: [
              new ImageRun({
                type: 'png',
                data: buf1,
                transformation: { width: 280, height: 210 },
              }),
            ],
          })
        );
        if (detail_1) {
          cell1Children.push(
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 120 },
              children: [new TextRun({ text: detail_1, size: 24, font: 'Times New Roman' })],
            })
          );
        }
      }
    }

    if (img2) {
      const isObj2 = typeof img2 === 'object' && img2 !== null;
      const base64_2 = isObj2 ? (img2 as AppendixImage).base64 : (img2 as string);
      const detail_2 = isObj2 ? (img2 as AppendixImage).detail : '';
      const buf2 = parseBase64Image(base64_2);

      if (buf2) {
        cell2Children.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 60 },
            children: [
              new ImageRun({
                type: 'png',
                data: buf2,
                transformation: { width: 280, height: 210 },
              }),
            ],
          })
        );
        if (detail_2) {
          cell2Children.push(
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 120 },
              children: [new TextRun({ text: detail_2, size: 24, font: 'Times New Roman' })],
            })
          );
        }
      }
    }

    rows.push(
      new TableRow({
        children: [
          new TableCell({ children: cell1Children, width: { size: 50, type: WidthType.PERCENTAGE } }),
          new TableCell({ children: cell2Children, width: { size: 50, type: WidthType.PERCENTAGE } }),
        ],
      })
    );
  }

  return new Table({
    rows,
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: tableBorders,
  });
}

function buildFeaturedImageLayout(imgData: string | AppendixImage) {
  const children: any[] = [];
  const isObj = typeof imgData === 'object' && imgData !== null;
  const base64 = isObj ? (imgData as AppendixImage).base64 : (imgData as string);
  const detail = isObj ? (imgData as AppendixImage).detail : '';
  const buf = parseBase64Image(base64);

  if (buf) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 100 },
        children: [
          new ImageRun({
            type: 'png',
            data: buf,
            transformation: { width: 420, height: 315 },
          }),
        ],
      })
    );

    if (detail) {
      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [new TextRun({ text: detail, size: 24, font: 'Times New Roman' })],
        })
      );
    }
  }

  return children;
}

function buildReportFooterBlock(footer?: ReportFooterData): any[] {
  const preparedByLabel = footer?.preparedByLabel?.trim() || 'Prepared By:';
  const preparedByName = footer?.preparedByName?.trim() || 'HAYNA G. DAUD';
  const checkedByLabel = footer?.checkedByLabel?.trim() || 'Checked By:';
  const checkedByName = footer?.checkedByName?.trim() || 'ANDRES S. TAPALES JR.';
  const checkedByRole = footer?.checkedByRole?.trim() || 'Supervising Statistical Specialist';
  const dateLabel = footer?.dateLabel?.trim() || 'Date:';
  const dateValue = footer?.dateValue?.trim() || 'August 15, 2026';
  const officeInCharge = footer?.officeInCharge?.trim() || 'Office-in-Charge';

  const footerTable = new Table({
    width: { size: 94, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      insideHorizontal: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      insideVertical: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 10, bottom: 10, left: 10, right: 10 },
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                spacing: { after: 12, line: 276, lineRule: 'auto' },
                children: [new TextRun({ text: preparedByLabel, bold: true, size: 24, font: 'Times New Roman' })],
              })
            ],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 10, bottom: 10, left: 10, right: 10 },
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                spacing: { after: 12, line: 276, lineRule: 'auto' },
                children: [new TextRun({ text: checkedByLabel, bold: true, size: 24, font: 'Times New Roman' })],
              })
            ],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 8, bottom: 8, left: 10, right: 10 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 6, after: 10, line: 276, lineRule: 'auto' },
                children: [new TextRun({ text: preparedByName, bold: true, size: 24, font: 'Times New Roman' })],
              })
            ],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 8, bottom: 8, left: 10, right: 10 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 6, after: 10, line: 276, lineRule: 'auto' },
                children: [new TextRun({ text: checkedByName, bold: true, size: 24, font: 'Times New Roman' })],
              })
            ],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 10, bottom: 10, left: 10, right: 10 },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 8, after: 8, line: 276, lineRule: 'auto' }, children: [new TextRun({ text: '', size: 24, font: 'Times New Roman' })] })],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 8, bottom: 8, left: 10, right: 10 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 8, after: 10, line: 276, lineRule: 'auto' },
                children: [new TextRun({ text: checkedByRole, size: 24, font: 'Times New Roman' })],
              })
            ],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 8, bottom: 10, left: 10, right: 10 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 8, after: 8, line: 276, lineRule: 'auto' },
                children: [
                  new TextRun({ text: `${dateLabel} `, bold: true, size: 24, font: 'Times New Roman' }),
                  new TextRun({ text: dateValue, bold: true, size: 24, font: 'Times New Roman' }),
                ],
              })
            ],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 8, bottom: 10, left: 10, right: 10 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 8, after: 8, line: 276, lineRule: 'auto' },
                children: [new TextRun({ text: officeInCharge, size: 24, font: 'Times New Roman' })],
              })
            ],
          }),
        ],
      }),
    ],
  });

  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0, line: 276, lineRule: 'auto' },
      children: [footerTable],
    }),
  ];
}

function buildDailyJournalAppendixPage(appendicesData: AppendicesData) {
  const children: any[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 360 },
      children: [new TextRun({ text: 'APPENDICES', bold: true, size: 24, font: 'Times New Roman' })],
    }),
  ];

  if (!appendicesData || !appendicesData.dailyJournal || !Array.isArray(appendicesData.dailyJournal)) return children;

  const layout = appendicesData.dailyJournalLayout === 'no-photo'
    ? 'no-photo'
    : appendicesData.dailyJournalLayout === 'report' ? 'report' : 'current';

  if (layout === 'no-photo') {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 180, after: 180 },
        children: [new TextRun({ text: 'DAILY JOURNAL', bold: true, size: 24, font: 'Times New Roman' })],
      })
    );

    appendicesData.dailyJournal.forEach((weekData) => {
      const totalHours = weekData.activities.reduce((total, activity) => total + (Number(activity.hours) || 0), 0);
      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 120, after: 60 },
          keepNext: true,
          children: [new TextRun({ text: `WEEK ${weekData.weekNumber}`, bold: true, size: 24, font: 'Times New Roman' })],
        }),
        buildCurrentWeeklyTable(weekData.activities, totalHours),
      );
    });

    return children;
  }

  if (layout === 'current') {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 180, after: 90 },
        children: [new TextRun({ text: 'WEEKLY WORK ACTIVITIES', bold: true, size: 24, font: 'Times New Roman' })],
      })
    );
  }

  const narrativeLineSpacing = 360;

  appendicesData.dailyJournal.forEach((weekData: DailyJournalWeek, index: number) => {
    if (layout === 'report') {
      // Restore the original daily layout behavior, which keeps documentation inside the table.
      const reportImages = weekData.reportLayoutImages || weekData.images || [];
      const weekParagraph = new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 120 },
        pageBreakBefore: index > 0,
        children: [new TextRun({ text: `WEEK ${weekData.weekNumber}`, bold: true, size: 24, font: 'Times New Roman' })],
      });
      children.push(
        weekParagraph,
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 60, after: 120 },
          children: [new TextRun({ text: 'WEEKLY ACCOMPLISHMENT REPORT', bold: true, size: 24, font: 'Times New Roman' })],
        }),
        buildDailyReportTable(weekData.activities, reportImages),
      );

      return;
    }

    // Use current layout images, fall back to shared images for backwards compatibility
    const currentImages = weekData.currentLayoutImages || weekData.images || [];
    const currentWeekParagraph = new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 80 },
      pageBreakBefore: index > 0,
      children: [new TextRun({ text: `WEEK ${weekData.weekNumber}`, bold: true, size: 24, font: 'Times New Roman' })],
    });
    children.push(
      currentWeekParagraph,
      buildCurrentWeeklyTable(weekData.activities, weekData.totalHours),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 120 },
        children: [new TextRun({ text: `Table ${weekData.weekNumber}: Week ${weekData.weekNumber}`, bold: true, size: 24, font: 'Times New Roman' })],
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { before: 0, after: 0, line: narrativeLineSpacing, lineRule: 'auto' },
        children: [new TextRun({ text: weekData.narrative || 'No narrative provided.', size: 24, font: 'Times New Roman' })],
      })
    );

    if (currentImages.length > 0) {
      children.push(buildImageGridTable(currentImages));
    }
  });

  return children;
}

function buildOtherAppendixPages(appendicesData: AppendicesData) {
  const children: any[] = [];
  if (!appendicesData) return children;

  const appendixList = [
    { title: 'Certificate of Participation (PRIME Seminar)', key: 'certParticipation' as const },
    { title: 'One page narrative report of PRIME', key: 'primeNarrative' as const, isText: true },
    { title: 'Resume and Application Letter', key: 'resume' as const },
    { title: 'Evaluation of Grades & Validated Enrolment Form', key: 'grades' as const },
    { title: 'Medical Certificate & Parent\'s Waiver', key: 'medicalWaiver' as const },
    { title: 'Letter of Acceptance', key: 'letterAcceptance' as const },
    { title: 'Daily Time Record (DTR) - hours rendered', key: 'dtr' as const },
    { title: 'Student’s Intern Rating Sheet & Performance Evaluation', key: 'ratingSheet' as const },
    { title: 'Certificate of Completion & Memorandum of Agreement', key: 'certCompletion' as const },
  ];

  appendixList.forEach(app => {
    const data = appendicesData[app.key];
    if (!data) return;

    if (children.length > 0) {
      children.push(new Paragraph({ pageBreakBefore: true }));
    }
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 240, after: 240 },
        children: [new TextRun({ text: app.title, bold: true, size: 24, font: 'Times New Roman' })],
      })
    );

    if (app.isText) {
      if (app.key === 'primeNarrative') {
        const primeImages = (appendicesData.primeReportImages || []).slice(0, 2);
        if (primeImages.length > 0) {
          children.push(buildImageGridTable(primeImages));
        }
        children.push(
          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { before: 0, after: 0, line: 360, lineRule: 'auto' },
            children: [new TextRun({ text: String(data), size: 24, font: 'Times New Roman' })],
          })
        );
        return;
      }

      children.push(
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun({ text: data, size: 24, font: 'Times New Roman' })],
        })
      );
    } else {
      children.push(...buildFeaturedImageLayout(data));
    }
  });

  return children;
}

function buildAppendicesPage(appendicesData: AppendicesData) {
  return [
    ...buildDailyJournalAppendixPage(appendicesData),
    ...buildOtherAppendixPages(appendicesData),
  ];
}

function buildCurrentWeeklyTable(activities: DailyActivity[], totalHours: number | string) {
  const TABLE_TEXT_SIZE = 24;
  const tableRows = [
    new TableRow({
      cantSplit: true,
      children: [
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 28, after: 28 }, children: [new TextRun({ text: 'DAY', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 16, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 28, after: 28 }, children: [new TextRun({ text: 'DATE', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 18, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 28, after: 28 }, children: [new TextRun({ text: 'ACCOMPLISHMENT', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 52, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 28, after: 28 }, children: [new TextRun({ text: 'NO. OF WORKING HOURS', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 14, type: WidthType.PERCENTAGE } }),
      ],
    }),
  ];

  const rows = activities && activities.length > 0 ? activities : [{ day: 'Day 1', date: '', accomplishment: '', hours: '' }];

  rows.forEach((act) => {
    tableRows.push(
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({
            width: { size: 16, type: WidthType.PERCENTAGE },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 18, after: 18 }, children: [new TextRun({ text: act.day || '', size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })],
            verticalAlign: 'center',
          }),
          new TableCell({
            width: { size: 18, type: WidthType.PERCENTAGE },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 18, after: 18 }, children: [new TextRun({ text: act.date || '', size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })],
            verticalAlign: 'center',
          }),
          new TableCell({
            width: { size: 52, type: WidthType.PERCENTAGE },
            children: (() => {
              const accomplishmentText = act.accomplishment || 'No accomplishment added.';
              const lines = accomplishmentText.split('\n').filter(line => line.trim());
              if (lines.length === 0) {
                return [new Paragraph({ alignment: AlignmentType.LEFT, spacing: { before: 18, after: 18 }, children: [new TextRun({ text: 'No accomplishment added.', size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })];
              }
              return lines.map((line, idx) =>
                new Paragraph({
                  alignment: AlignmentType.LEFT,
                  spacing: { before: 14, after: idx === lines.length - 1 ? 14 : 8 },
                  children: [new TextRun({ text: line.trim(), size: TABLE_TEXT_SIZE, font: 'Times New Roman' })]
                })
              );
            })(),
          }),
          new TableCell({
            width: { size: 14, type: WidthType.PERCENTAGE },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 18, after: 18 }, children: [new TextRun({ text: String(act.hours ?? ''), size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })],
            verticalAlign: 'center',
          }),
        ],
      })
    );
  });

  tableRows.push(
    new TableRow({
      cantSplit: true,
      children: [
        new TableCell({ 
          columnSpan: 3, 
          children: [new Paragraph({ alignment: AlignmentType.LEFT, spacing: { before: 24, after: 24 }, children: [new TextRun({ text: 'TOTAL NUMBER OF HOURS', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })] 
        }),
        new TableCell({ 
          children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 24, after: 24 }, children: [new TextRun({ text: String(totalHours), bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })] 
        }),
      ],
    })
  );

  return new Table({
    rows: tableRows,
    alignment: AlignmentType.CENTER,
    width: { size: 96, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    columnWidths: [1600, 1800, 5200, 1400],
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
    },
  });
}

function buildDailyReportTable(activities: DailyActivity[], images: (string | AppendixImage)[] = []) {
  const TABLE_TEXT_SIZE = 24;
  const tableRows = [
    new TableRow({
      cantSplit: true,
      children: [
        new TableCell({ margins: { top: 18, bottom: 18, left: 24, right: 24 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: 'DATE', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 18, type: WidthType.PERCENTAGE } }),
        new TableCell({ margins: { top: 18, bottom: 18, left: 24, right: 24 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: 'ACCOMPLISHMENT', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 34, type: WidthType.PERCENTAGE } }),
        new TableCell({ margins: { top: 18, bottom: 18, left: 24, right: 24 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: 'HOURS', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 10, type: WidthType.PERCENTAGE } }),
        new TableCell({ margins: { top: 18, bottom: 18, left: 24, right: 24 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: 'DOCUMENTATION', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 38, type: WidthType.PERCENTAGE } }),
      ],
    }),
  ];

  const rows = activities && activities.length > 0 ? activities : [{ day: 'Day 1', date: '', accomplishment: '', hours: '' }];

  rows.forEach((act, index) => {
    const imageData = images[index] || null;
    const imageCellChildren: any[] = [];

    if (imageData) {
      const isObj = typeof imageData === 'object' && imageData !== null;
      const base64 = isObj ? (imageData as AppendixImage).base64 : (imageData as string);
      const detail = isObj ? (imageData as AppendixImage).detail : '';
      const buf = parseBase64Image(base64);

      if (buf) {
        imageCellChildren.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 0 },
            children: [new ImageRun({ type: 'png', data: buf, transformation: { width: 260, height: 170 } })],
          })
        );

        if (detail) {
          imageCellChildren.push(
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 0, after: 0 },
              children: [new TextRun({ text: detail, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })],
            })
          );
        }
      }
    }

    tableRows.push(
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({ margins: { top: 12, bottom: 12, left: 18, right: 18 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: act.date || act.day || '', size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], verticalAlign: 'center' }),
          new TableCell({
            margins: { top: 12, bottom: 12, left: 18, right: 18 },
            children: (() => {
              const accomplishmentText = act.accomplishment || 'No accomplishment added.';
              const lines = accomplishmentText.split('\n').filter(line => line.trim());
              if (lines.length === 0) {
                return [new Paragraph({ alignment: AlignmentType.LEFT, spacing: { before: 0, after: 0 }, bullet: { level: 0 }, children: [new TextRun({ text: 'No accomplishment added.', size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })];
              }
              return lines.map((line, idx) =>
                new Paragraph({
                  alignment: AlignmentType.LEFT,
                  spacing: { before: 0, after: idx === lines.length - 1 ? 0 : 0 },
                  bullet: { level: 0 },
                  children: [new TextRun({ text: line.trim(), size: TABLE_TEXT_SIZE, font: 'Times New Roman' })]
                })
              );
            })()
          }),
          new TableCell({ margins: { top: 12, bottom: 12, left: 18, right: 18 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: String(act.hours ?? ''), size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], verticalAlign: 'center' }),
          new TableCell({ margins: { top: 12, bottom: 12, left: 18, right: 18 }, children: imageCellChildren.length > 0 ? imageCellChildren : [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '', size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })] }),
        ],
      })
    );
  });

  return new Table({
    rows: tableRows,
    alignment: AlignmentType.CENTER,
    width: { size: 94, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    columnWidths: [1700, 3300, 1000, 4000],
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
    },
  });
}

function buildWeeklyReportTable(activities: DailyActivity[], images: (string | AppendixImage)[] = []) {
  const TABLE_TEXT_SIZE = 24;
  const tableRows = [
    new TableRow({
      cantSplit: true,
      children: [
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'DAY', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 16, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'DATE', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 18, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'ACCOMPLISHMENT', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 52, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'NO. OF WORKING HOURS', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], width: { size: 14, type: WidthType.PERCENTAGE } }),
      ],
    }),
  ];

  const rows = activities && activities.length > 0 ? activities : [{ day: 'Day 1', date: '', accomplishment: '', hours: '' }];

  rows.forEach((act) => {
    tableRows.push(
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: act.day || '', size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], verticalAlign: 'center' }),
          new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: act.date || '', size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })], verticalAlign: 'center' }),
          new TableCell({ 
            children: (() => {
              const accomplishmentText = act.accomplishment || 'No accomplishment added.';
              const lines = accomplishmentText.split('\n').filter(line => line.trim());
              if (lines.length === 0) {
                return [new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 40 }, children: [new TextRun({ text: 'No accomplishment added.', size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })];
              }
              return lines.map((line, idx) => 
                new Paragraph({ 
                  alignment: AlignmentType.LEFT, 
                  spacing: { after: idx === lines.length - 1 ? 40 : 20 },
                  children: [new TextRun({ text: line.trim(), size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] 
                })
              );
            })()
          }),
          new TableCell({
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(act.hours ?? ''), size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })],
            verticalAlign: 'center',
          }),
        ],
      })
    );
  });

  const totalCellValue = rows.reduce((sum, row) => sum + Number(row.hours || 0), 0);

  tableRows.push(
    new TableRow({
      cantSplit: true,
      children: [
        new TableCell({
          columnSpan: 3,
          children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: 'TOTAL NUMBER OF HOURS', bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })],
        }),
        new TableCell({
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(totalCellValue), bold: true, size: TABLE_TEXT_SIZE, font: 'Times New Roman' })] })],
        }),
      ],
    })
  );

  return new Table({
    rows: tableRows,
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
    },
  });
}
