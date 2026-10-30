import { NextResponse } from 'next/server';

export const runtime = 'nodejs';

const exhaustedUntil = new Map<string, number>();
let providerCursor = 0;
const keyCursor = new Map<string, number>();

function getKeys(primary: string | undefined, list: string | undefined): string[] {
  const raw = [list || '', primary || ''].join('\n');
  return Array.from(new Set(
    raw
      .split(/[,\n;]/)
      .map((key) => key.trim())
      .filter(Boolean),
  ));
}

function isRateLimit(errorText: string, status: number): boolean {
  return status === 429 || /rate.?limit|quota|resource.?exhausted|too many requests|limit exceeded/i.test(errorText);
}

function availableKey(keys: string[], provider: string): string | undefined {
  const now = Date.now();
  const available: string[] = [];

  for (const key of keys) {
    const mapKey = `${provider}:${key}`;
    const until = exhaustedUntil.get(mapKey) || 0;
    if (until <= now) {
      exhaustedUntil.delete(mapKey);
      available.push(key);
    }
  }

  if (!available.length) return undefined;
  const cursor = keyCursor.get(provider) || 0;
  const selected = available[cursor % available.length];
  keyCursor.set(provider, (cursor + 1) % available.length);
  return selected;
}

function prepareAIContext(value: unknown): string {
  const seen = new WeakSet();

  const sanitize = (item: unknown): unknown => {
    if (typeof item === 'string') {
      return item.startsWith('data:image/') ? '[image attachment omitted]' : item;
    }
    if (!item || typeof item !== 'object') return item;
    if (seen.has(item as object)) return '[circular reference omitted]';
    seen.add(item as object);

    if (Array.isArray(item)) {
      return item.map((entry) => sanitize(entry));
    }

    return Object.fromEntries(
      Object.entries(item as Record<string, unknown>)
        .filter(([key, entry]) => {
          if (/base64|image|photo|picture|resume|medical|grade|dtr|certificate|letter|rating/i.test(key)) {
            if (typeof entry === 'string' && entry.startsWith('data:image/')) return false;
            if (Array.isArray(entry) && entry.length > 0) return false;
          }
          return true;
        })
        .map(([key, entry]) => [key, sanitize(entry)])
    );
  };

  return JSON.stringify(sanitize(value)).slice(0, 50000);
}

export async function POST(request: Request) {
  const groqKeys = getKeys(process.env.GROQ_API_KEY, process.env.GROQ_API_KEYS);
  const geminiKeys = getKeys(process.env.GEMINI_API_KEY, process.env.GEMINI_API_KEYS);
  if (!groqKeys.length && !geminiKeys.length) {
    return NextResponse.json(
      { error: 'Neither GROQ_API_KEY nor GEMINI_API_KEY is configured on the server.' },
      { status: 500 },
    );
  }

  try {
    const body = await request.json();
    const title = String(body.title || 'Report section').slice(0, 200);
    const context = body.context && typeof body.context === 'object'
      ? prepareAIContext(body.context)
      : String(body.context || '').slice(0, 50000);
    const currentText = String(body.currentText || '').slice(0, 12000);
    const instructions = String(body.instructions || '').slice(0, 2000);

    const prompt = [
      'You are a professional academic writing assistant for an On-the-Job Training (OJT) narrative report.',
      `Section: ${title}`,
      'Use the complete report data supplied below as the source of truth: the organization details, student and placement information, daily journal activities, weekly narratives, SWOT data, assigned tasks, case-analysis details, and all other relevant report fields.',
      'Analyze relationships across the entire report so the generated section is consistent with the student, organization, OJT placement, dates, accomplishments, and professional context.',
      'For acknowledgements and organization-specific sections, identify the exact active placement from the report data (for example, a named office such as MISO) and refer to it by its exact name. Do not replace it with a generic term such as "the front page," university, office, or organization unless that exact term is present elsewhere.',
      'Write natural, humanized, formal, and specific academic prose suitable for direct insertion into a Word document. Use varied sentence structure and a confident, professional tone.',
      'Do not make the text generic, repetitive, promotional, or obviously AI-generated. Do not invent names, organizations, dates, duties, achievements, skills, statistics, or results.',
      'Preserve facts exactly when they are explicit. Connect only facts that are logically supported by the supplied data. Prefer concrete workplace details over vague language.',
      'Use first-person language when the section naturally describes the student experience. Otherwise use a formal academic or organizational voice.',
      'Do not use markdown, bullet symbols, headings, labels, or brackets.',
      'Separate paragraphs with one blank line.',
      instructions ? `Additional instructions: ${instructions}` : '',
      currentText ? `Current text to improve: ${currentText}` : '',
      `Report data: ${context}`,
    ].filter(Boolean).join('\n\n');

    let text = '';
    const providerErrors: string[] = [];
    let provider = '';

    const providers = providerCursor % 2 === 0
      ? ['groq', 'gemini']
      : ['gemini', 'groq'];
    providerCursor += 1;

    for (const providerName of providers) {
      if (text) break;

      const keys = providerName === 'groq' ? groqKeys : geminiKeys;
      for (let rotation = 0; rotation < keys.length && !text; rotation += 1) {
        const apiKey = availableKey(keys, providerName);
        if (!apiKey) break;

        try {
          if (providerName === 'groq') {
            const response = await fetch('https://api.groq.com/openai/v1/chat/completions', {
              method: 'POST',
              headers: {
                'Content-Type': 'application/json',
                Authorization: 'Bearer ' + apiKey,
              },
              body: JSON.stringify({
                model: process.env.GROQ_MODEL || 'openai/gpt-oss-20b',
                messages: [{ role: 'user', content: prompt }],
                temperature: 0.35,
                max_tokens: 2200,
              }),
            });
            const result = await response.json();
            if (response.ok) {
              text = String(result?.choices?.[0]?.message?.content || '').trim();
              if (text) provider = 'groq';
            } else {
              const message = String(result?.error?.message || 'Groq could not generate the text.');
              providerErrors.push(`Groq: ${message}`);
              if (isRateLimit(message, response.status)) {
                exhaustedUntil.set(`groq:${apiKey}`, Date.now() + 60_000);
              }
            }
          } else {
            const geminiModel = process.env.GEMINI_MODEL || 'gemini-flash-latest';
            const response = await fetch(
              `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(geminiModel)}:generateContent`,
              {
                method: 'POST',
                headers: {
                  'Content-Type': 'application/json',
                  'X-goog-api-key': apiKey,
                },
                body: JSON.stringify({
                  contents: [{ parts: [{ text: prompt }] }],
                  generationConfig: {
                    temperature: 0.35,
                    maxOutputTokens: 2200,
                  },
                }),
              },
            );
            const result = await response.json();
            if (response.ok) {
              text = String(
                result?.candidates?.[0]?.content?.parts
                  ?.map((part: { text?: string }) => part.text || '')
                  .join('')
                  .trim() || '',
              );
              if (text) provider = 'gemini';
            } else {
              const message = String(result?.error?.message || 'Gemini could not generate the text.');
              providerErrors.push(`Gemini: ${message}`);
              if (isRateLimit(message, response.status)) {
                exhaustedUntil.set(`gemini:${apiKey}`, Date.now() + 60_000);
              }
            }
          }
        } catch (error) {
          providerErrors.push(
            `${providerName === 'groq' ? 'Groq' : 'Gemini'}: ${
              error instanceof Error ? error.message : 'Provider request failed.'
            }`,
          );
        }
      }
    }

    if (!text) {
      return NextResponse.json(
        { error: providerErrors.join(' | ') || 'The AI provider returned no text.' },
        { status: 502 },
      );
    }

    return NextResponse.json({ text, provider });
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Could not generate text.';
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
