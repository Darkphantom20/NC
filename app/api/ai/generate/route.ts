import { NextResponse } from 'next/server';

export const runtime = 'nodejs';

export async function POST(request: Request) {
  const apiKey = process.env.GROQ_API_KEY;
  const geminiApiKey = process.env.GEMINI_API_KEY;
  if (!apiKey && !geminiApiKey) {
    return NextResponse.json(
      { error: 'Neither GROQ_API_KEY nor GEMINI_API_KEY is configured on the server.' },
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

    let text = '';
    let groqError = '';
    let provider = '';

    if (apiKey) {
      try {
        const groqResponse = await fetch('https://api.groq.com/openai/v1/chat/completions', {
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

        const groqResult = await groqResponse.json();
        if (groqResponse.ok) {
          text = String(groqResult?.choices?.[0]?.message?.content || '').trim();
          if (text) provider = 'groq';
        } else {
          groqError = String(
            groqResult?.error?.message || 'Groq could not generate the text.',
          );
        }
      } catch (error) {
        groqError = error instanceof Error ? error.message : 'Groq request failed.';
      }
    }

    if (!text && geminiApiKey) {
      const geminiModel = process.env.GEMINI_MODEL || 'gemini-flash-latest';
      const geminiResponse = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(geminiModel)}:generateContent`,
        {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'X-goog-api-key': geminiApiKey,
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

      const geminiResult = await geminiResponse.json();
      if (geminiResponse.ok) {
        text = String(
          geminiResult?.candidates?.[0]?.content?.parts
            ?.map((part: { text?: string }) => part.text || '')
            .join('')
            .trim() || '',
        );
        if (text) provider = 'gemini';
      } else {
        const geminiError = String(
          geminiResult?.error?.message || 'Gemini could not generate the text.',
        );
        return NextResponse.json(
          { error: groqError ? `${groqError} Gemini fallback also failed: ${geminiError}` : geminiError },
          { status: geminiResponse.status },
        );
      }
    }

    if (!text) {
      return NextResponse.json(
        { error: groqError || 'The AI provider returned no text.' },
        { status: 502 },
      );
    }

    return NextResponse.json({ text, provider });
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Could not generate text.';
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
