import { NextResponse } from 'next/server';

export const runtime = 'nodejs';

const NO_CONTEXT_KEYS = /(base64|dataurl|preview|image|photo|blob|signature)/i;

function sanitizeContext(value: unknown, depth = 0): unknown {
  if (value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return typeof value === 'string' ? value.slice(0, 3000) : value;
  }
  if (depth >= 5) return undefined;
  if (Array.isArray(value)) {
    return value
      .slice(0, 40)
      .map((item) => sanitizeContext(item, depth + 1))
      .filter((item) => item !== undefined);
  }
  if (typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value as Record<string, unknown>)
        .filter(([key]) => !NO_CONTEXT_KEYS.test(key))
        .map(([key, item]) => [key, sanitizeContext(item, depth + 1)])
        .filter(([, item]) => item !== undefined),
    );
  }
  return undefined;
}

export async function POST(request: Request) {
  const apiKey = process.env.GROQ_API_KEY;
  if (!apiKey) {
    return NextResponse.json(
      { error: 'GROQ_API_KEY is not configured on the server.' },
      { status: 500 },
    );
  }

  try {
    const body = await request.json();
    const title = String(body.title || 'Report section').slice(0, 200);
    const context = body.context && typeof body.context === 'object'
      ? JSON.stringify(sanitizeContext(body.context) || {}).slice(0, 12000)
      : String(body.context || '').slice(0, 20000);
    const currentText = String(body.currentText || '').slice(0, 12000);
    const instructions = String(body.instructions || '').slice(0, 2000);

    const prompt = [
      'Write a polished academic narrative for an OJT narrative report.',
      `Section: ${title}`,
      'Use the supplied report data only.',
      'Treat every person name, organization name, location, title, task, date, hour, skill, issue, and outcome as a fact that must be preserved exactly.',
      'If a required fact is missing from the supplied data, omit that fact rather than guessing or filling it in.',
      'Do not infer tools, responsibilities, qualifications, achievements, dates, locations, or outcomes that are not explicitly stated.',
      'Write complete, coherent paragraphs suitable for direct insertion into a Word document.',
      'Use first-person language only when appropriate and do not invent facts.',
      'Do not use markdown, bullet symbols, headings, labels, or brackets.',
      'Separate paragraphs with one blank line.',
      instructions ? `Additional instructions: ${instructions}` : '',
      currentText ? `Current text to improve: ${currentText}` : '',
      `Report data: ${context}`,
    ].filter(Boolean).join('\n\n');

    const response = await fetch('https://api.groq.com/openai/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: 'Bearer ' + apiKey,
      },
      body: JSON.stringify({
        model: process.env.GROQ_MODEL || 'openai/gpt-oss-20b',
        messages: [{ role: 'user', content: prompt }],
        temperature: 0.4,
        max_tokens: 1000,
      }),
    });

    const result = await response.json();
    if (!response.ok) {
      return NextResponse.json(
        { error: result?.error?.message || 'Groq could not generate the text.' },
        { status: response.status },
      );
    }

    const text = result?.choices?.[0]?.message?.content?.trim();
    if (!text) {
      return NextResponse.json({ error: 'Groq returned no text.' }, { status: 502 });
    }

    return NextResponse.json({ text });
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Could not generate text.';
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
