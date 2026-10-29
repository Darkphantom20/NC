import { NextResponse } from 'next/server';

export const runtime = 'nodejs';

export async function POST(request: Request) {
  const apiKey = process.env.GEMINI_API_KEY || process.env.GOOGLE_API_KEY;
  if (!apiKey) {
    return NextResponse.json(
      { error: 'GEMINI_API_KEY is not configured on the server.' },
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
      'Use first-person language when appropriate, but do not invent facts.',
      'Do not use markdown, bullet symbols, headings, labels, or brackets.',
      'Separate different paragraphs with one blank line.',
      instructions ? `Additional instructions: ${instructions}` : '',
      currentText ? `Current text (improve and align it if useful): ${currentText}` : '',
      `Report data: ${context}`,
    ].filter(Boolean).join('\n\n');

    const response = await fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent',
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-goog-api-key': apiKey,
        },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: {
            temperature: 0.4,
            maxOutputTokens: 1500,
          },
        }),
      },
    );

    const result = await response.json();
    if (!response.ok) {
      return NextResponse.json(
        { error: result?.error?.message || 'Gemini could not generate the text.' },
        { status: response.status },
      );
    }

    const text = result?.candidates?.[0]?.content?.parts
      ?.map((part: { text?: string }) => part.text || '')
      .join('')
      .trim();

    if (!text) {
      return NextResponse.json({ error: 'Gemini returned no text.' }, { status: 502 });
    }

    return NextResponse.json({ text });
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Could not generate text.';
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
