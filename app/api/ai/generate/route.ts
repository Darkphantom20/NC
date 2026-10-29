import { NextResponse } from 'next/server';

export const runtime = 'nodejs';

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
            temperature: 0.35,
            max_tokens: 2200,
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
              temperature: 0.35,
              maxOutputTokens: 2200,
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
