import { NextResponse } from 'next/server';

export const runtime = 'nodejs';

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
      ? JSON.stringify(body.context).slice(0, 20000)
      : String(body.context || '').slice(0, 20000);
    const currentText = String(body.currentText || '').slice(0, 12000);
    const instructions = String(body.instructions || '').slice(0, 2000);

    const prompt = [
      'Write a polished academic narrative for an OJT narrative report.',
      `Section: ${title}`,
      'Use the supplied report data only.',
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
        max_tokens: 1500,
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
