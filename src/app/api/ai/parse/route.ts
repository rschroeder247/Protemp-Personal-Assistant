import { NextRequest, NextResponse } from 'next/server';

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const {
      transcript,
      contacts = [],
      provider = 'gemini',
      apiKey,
      todayISO = new Date().toISOString().split('T')[0],
    } = body;

    if (!transcript) {
      return NextResponse.json({ actions: [] });
    }

    const effectiveKey =
      apiKey ||
      (provider === 'gemini'
        ? process.env.GEMINI_API_KEY
        : process.env.OPENAI_API_KEY);

    if (!effectiveKey) {
      return NextResponse.json(
        { error: 'No AI API key configured' },
        { status: 400 }
      );
    }

    const systemPrompt = `You are an executive Personal Assistant AI for Roland at Protemp.
Today's date is ${todayISO}.
Known contacts directory: ${JSON.stringify(contacts)}.

Extract all distinct actionable items from the user's voice note transcript and return ONLY valid JSON in this exact format:
{
  "actions": [
    {
      "action_type": "whatsapp" | "outlook_task" | "planner_task" | "outlook_email" | "outlook_event",
      "title": "Short title or subject",
      "description": "Polished message text (for WhatsApp/Email) or task details",
      "recipient_name": "Contact name if applicable",
      "recipient_phone": "Matched phone number if found in contacts",
      "recipient_email": "Matched email if found in contacts",
      "due_date": "YYYY-MM-DD if a due date or event date is mentioned",
      "start_time": "HH:MM (24h) if a time is mentioned",
      "priority": "low" | "normal" | "high" | "urgent"
    }
  ]
}
Rules:
- If the user says "WhatsApp" or "What app" or "message", set action_type to "whatsapp" and make description the exact polished message to send to that person (without "tell him that").
- If the user mentions "Planner", set action_type to "planner_task".
- If the user mentions "Outlook task", "task", "remind me", or "to do", set action_type to "outlook_task".
- If the user mentions "email" or "mail", set action_type to "outlook_email".
- If the user mentions "meeting", "schedule", or "calendar", set action_type to "outlook_event".`;

    if (provider === 'gemini') {
      const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${encodeURIComponent(
        effectiveKey.trim()
      )}`;

      const res = await fetch(geminiUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [
            {
              role: 'user',
              parts: [
                {
                  text: `${systemPrompt}\n\nVoice Note Transcript:\n"${transcript}"`,
                },
              ],
            },
          ],
          generationConfig: {
            responseMimeType: 'application/json',
            temperature: 0.1,
          },
        }),
      });

      const data = await res.json();
      const rawText =
        data?.candidates?.[0]?.content?.parts?.[0]?.text || '{"actions":[]}';
      const parsed = JSON.parse(rawText);
      return NextResponse.json({ actions: parsed.actions || [] });
    } else {
      const openAiRes = await fetch(
        'https://api.openai.com/v1/chat/completions',
        {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${effectiveKey.trim()}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            model: 'gpt-4o-mini',
            response_format: { type: 'json_object' },
            messages: [
              { role: 'system', content: systemPrompt },
              { role: 'user', content: transcript },
            ],
            temperature: 0.1,
          }),
        }
      );

      const data = await openAiRes.json();
      const rawText =
        data?.choices?.[0]?.message?.content || '{"actions":[]}';
      const parsed = JSON.parse(rawText);
      return NextResponse.json({ actions: parsed.actions || [] });
    }
  } catch (error: any) {
    console.error('AI parse route error:', error);
    return NextResponse.json(
      { error: error?.message || 'Failed to parse with AI' },
      { status: 500 }
    );
  }
}
