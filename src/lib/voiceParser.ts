import {
  ActionPriority,
  ActionType,
  AssistantAction,
  AssistantSettings,
  Contact,
} from './types';

export type DraftActionInput = Omit<
  AssistantAction,
  'id' | 'created_at' | 'updated_at'
>;

const DAYS_OF_WEEK = [
  'sunday',
  'monday',
  'tuesday',
  'wednesday',
  'thursday',
  'friday',
  'saturday',
];

function formatDateISO(d: Date): string {
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

export function extractDateAndTime(text: string): {
  dueDate?: string;
  startTime?: string;
  cleanedText: string;
} {
  const lower = text.toLowerCase();
  const now = new Date();
  let dueDate: string | undefined;
  let startTime: string | undefined;
  let cleanedText = text;

  // 1. Check relative dates
  if (/\btoday\b|\btonight\b/i.test(lower)) {
    dueDate = formatDateISO(now);
  } else if (/\btomorrow\b/i.test(lower)) {
    const d = new Date(now);
    d.setDate(d.getDate() + 1);
    dueDate = formatDateISO(d);
  } else if (/\bnext week\b/i.test(lower)) {
    const d = new Date(now);
    d.setDate(d.getDate() + 7);
    dueDate = formatDateISO(d);
  } else {
    const inDaysMatch = lower.match(/\bin\s+(\d+)\s+days?\b/i);
    if (inDaysMatch) {
      const d = new Date(now);
      d.setDate(d.getDate() + parseInt(inDaysMatch[1], 10));
      dueDate = formatDateISO(d);
    } else {
      for (let i = 0; i < DAYS_OF_WEEK.length; i++) {
        const dayName = DAYS_OF_WEEK[i];
        const regex = new RegExp(`\\b(?:on\\s+|by\\s+|next\\s+|this\\s+)?${dayName}\\b`, 'i');
        if (regex.test(lower)) {
          const d = new Date(now);
          const currentDay = d.getDay();
          let diff = (i - currentDay + 7) % 7;
          if (diff === 0) diff = 7;
          d.setDate(d.getDate() + diff);
          dueDate = formatDateISO(d);
          break;
        }
      }
    }
  }

  // 2. Check time expressions (e.g. "at 2pm", "at 14:30", "at 9:00 am")
  const timeMatch = lower.match(
    /\b(?:at\s+)?(\d{1,2})(?::(\d{2}))?\s*(am|pm)\b|\bat\s+(\d{1,2}):(\d{2})\b/i
  );
  if (timeMatch) {
    if (timeMatch[1]) {
      let hours = parseInt(timeMatch[1], 10);
      const mins = timeMatch[2] ? parseInt(timeMatch[2], 10) : 0;
      const ampm = timeMatch[3]?.toLowerCase();
      if (ampm === 'pm' && hours < 12) hours += 12;
      if (ampm === 'am' && hours === 12) hours = 0;
      startTime = `${String(hours).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
    } else if (timeMatch[4]) {
      startTime = `${String(parseInt(timeMatch[4], 10)).padStart(2, '0')}:${timeMatch[5]}`;
    }
  } else if (/\bmorning\b/i.test(lower)) {
    startTime = '09:00';
  } else if (/\bafternoon\b/i.test(lower)) {
    startTime = '14:00';
  }

  return { dueDate, startTime, cleanedText };
}

export function detectPriority(text: string): ActionPriority {
  const lower = text.toLowerCase();
  if (/\b(urgent|urgently|asap|critical|emergency|immediately)\b/.test(lower)) {
    return 'urgent';
  }
  if (/\b(high priority|important|vital)\b/.test(lower)) {
    return 'high';
  }
  if (/\b(low priority|whenever|someday)\b/.test(lower)) {
    return 'low';
  }
  return 'normal';
}

export function findMatchingContact(
  nameCandidate: string | undefined,
  fullText: string,
  contacts: Contact[]
): Contact | undefined {
  if (!contacts || contacts.length === 0) return undefined;

  if (nameCandidate) {
    const cleanTarget = nameCandidate.trim().toLowerCase();
    const exactOrPartial = contacts.find(
      (c) =>
        c.name.toLowerCase() === cleanTarget ||
        c.name.toLowerCase().startsWith(cleanTarget) ||
        c.name.toLowerCase().includes(cleanTarget)
    );
    if (exactOrPartial) return exactOrPartial;
  }

  const lowerFull = fullText.toLowerCase();
  for (const c of contacts) {
    const firstName = c.name.split(/\s+/)[0]?.toLowerCase();
    if (
      c.name &&
      (lowerFull.includes(c.name.toLowerCase()) ||
        (firstName &&
          firstName.length >= 3 &&
          new RegExp(`\\b${firstName}\\b`, 'i').test(lowerFull)))
    ) {
      return c;
    }
  }

  return undefined;
}

/**
 * Splits a multi-command voice note into logical clauses.
 */
function splitIntoSegments(transcript: string): string[] {
  const normalized = transcript.trim();
  if (!normalized) return [];

  // Split when transition phrases introduce a new command verb
  const splitRegex =
    /\s+(?:and\s+then|and\s+also|also|then|next|plus|after\s+that|,?\s+and)\s+(?=(?:send|whatsapp|message|text|create|add|make|generate|schedule|book|email|mail|remind|set|put)\b)/i;

  const rawParts = normalized
    .split(splitRegex)
    .flatMap((part) =>
      // Also split on sentence boundaries if the next sentence starts with an action trigger
      part.split(
        /\.\s+(?=(?:send|whatsapp|message|text|create|add|make|generate|schedule|book|email|mail|remind|set|put)\b)/i
      )
    )
    .map((s) => s.trim())
    .filter(Boolean);

  return rawParts.length > 0 ? rawParts : [normalized];
}

function parseSingleClause(
  clause: string,
  contacts: Contact[],
  settings: AssistantSettings,
  voiceNoteId?: string
): DraftActionInput {
  const lower = clause.toLowerCase();
  const priority = detectPriority(clause);
  const { dueDate, startTime } = extractDateAndTime(clause);

  // 1. Detect Action Type
  let actionType: ActionType = 'outlook_task';

  if (
    /\b(whatsapp|whats\s*app|what's\s*app|what\s+app|wa\s+message|text\s+message)\b/i.test(
      lower
    ) ||
    /\b(?:send\s+a?\s*message|message|text)\s+to\b/i.test(lower)
  ) {
    actionType = 'whatsapp';
  } else if (/\b(planner|plan\s+task|team\s+task|bucket)\b/i.test(lower)) {
    actionType = 'planner_task';
  } else if (
    /\b(email|e-mail|send\s+an?\s+email|mail\s+to|draft\s+an?\s+email)\b/i.test(lower)
  ) {
    actionType = 'outlook_email';
  } else if (
    /\b(meeting|calendar|schedule\s+a\s+call|schedule\s+a\s+meeting|appointment|event)\b/i.test(
      lower
    )
  ) {
    actionType = 'outlook_event';
  } else {
    actionType = 'outlook_task';
  }

  // 2. Parse WhatsApp intent
  if (actionType === 'whatsapp') {
    let recipientName = '';
    let messageBody = clause;

    // Patterns like: "Send a WhatsApp to John saying I will be there at 2pm"
    // or "WhatsApp Sarah that the quote is ready"
    const waMatch =
      clause.match(
        /(?:send\s+(?:a\s+)?(?:whatsapp|whats\s*app|what's\s*app|what\s+app|message|text)\s+(?:message\s+)?to|whatsapp|whats\s*app|what's\s*app|what\s+app|message|text)\s+([A-Za-z0-9_]+(?:\s+[A-Za-z0-9_]+)?)\s+(?:saying\s+that|saying|telling\s+(?:him|her|them)\s+that|telling\s+(?:him|her|them)|to\s+say\s+that|to\s+say|that|:|-)\s+(.+)/i
      ) ||
      clause.match(
        /(?:send\s+(?:a\s+)?(?:whatsapp|whats\s*app|what's\s*app|what\s+app|message|text)\s+(?:message\s+)?to|whatsapp|whats\s*app|what's\s*app|what\s+app)\s+([A-Za-z]+)\s+(.+)/i
      );

    if (waMatch) {
      recipientName = waMatch[1].trim();
      messageBody = waMatch[2].trim();
    } else {
      // Strip leading command words
      messageBody = clause
        .replace(
          /^(?:please\s+)?(?:can\s+you\s+)?(?:send\s+(?:a\s+)?)?(?:whatsapp|whats\s*app|what's\s*app|what\s+app|message|text)\s*(?:message\s*)?(?:saying\s*)?/i,
          ''
        )
        .trim();
    }

    // Check if an explicit phone number was spoken in the clause
    const phoneInTextMatch = clause.match(/(?:\+?\d[\d\s-]{7,14}\d)/);
    const matchedContact = findMatchingContact(recipientName, clause, contacts);

    const finalRecipientName =
      matchedContact?.name ||
      (recipientName
        ? recipientName.charAt(0).toUpperCase() + recipientName.slice(1)
        : 'Contact');
    const finalPhone =
      matchedContact?.phone ||
      (phoneInTextMatch ? phoneInTextMatch[0].replace(/[\s-]/g, '') : '');

    // Capitalize first letter of message body
    const cleanMsg =
      messageBody.charAt(0).toUpperCase() + messageBody.slice(1);

    return {
      voice_note_id: voiceNoteId || null,
      action_type: 'whatsapp',
      title: `WhatsApp to ${finalRecipientName}`,
      description: cleanMsg,
      recipient_name: finalRecipientName,
      recipient_phone: finalPhone,
      priority,
      status: 'draft',
    };
  }

  // 3. Parse Outlook Email intent
  if (actionType === 'outlook_email') {
    let recipientName = '';
    let body = clause;
    let subject = 'Follow up from Roland';

    const emailMatch = clause.match(
      /(?:send\s+(?:an?\s+)?(?:email|e-mail|mail)\s+to|email|e-mail|mail)\s+([A-Za-z]+(?:\s+[A-Za-z]+)?)\s+(?:about|regarding|subject|saying\s+that|saying|that|:)\s+(.+)/i
    );

    if (emailMatch) {
      recipientName = emailMatch[1].trim();
      body = emailMatch[2].trim();
      subject =
        body.length > 55 ? `${body.slice(0, 52).trim()}...` : body;
    } else {
      body = clause
        .replace(
          /^(?:please\s+)?(?:send\s+(?:an?\s+)?|draft\s+(?:an?\s+)?|create\s+(?:an?\s+)?)?(?:outlook\s+)?(?:email|e-mail|mail)\s*(?:to\s+)?/i,
          ''
        )
        .trim();
      subject = body.length > 55 ? `${body.slice(0, 52).trim()}...` : body;
    }

    const matchedContact = findMatchingContact(recipientName, clause, contacts);

    return {
      voice_note_id: voiceNoteId || null,
      action_type: 'outlook_email',
      title: subject.charAt(0).toUpperCase() + subject.slice(1),
      description: body.charAt(0).toUpperCase() + body.slice(1),
      recipient_name: matchedContact?.name || recipientName || '',
      recipient_email: matchedContact?.email || '',
      priority,
      status: 'draft',
    };
  }

  // 4. Parse Planner Task intent
  if (actionType === 'planner_task') {
    const cleanTitle = clause
      .replace(
        /^(?:please\s+)?(?:can\s+you\s+)?(?:create|add|generate|make|put|set)\s+(?:a\s+|an\s+)?(?:new\s+)?(?:microsoft\s+)?(?:planner\s+task|task\s+in\s+planner|planner\s+item|plan\s+task)\s*(?:to\s+|for\s+|called\s+|named\s+|that\s+)?/i,
        ''
      )
      .replace(/\bin\s+planner\b/gi, '')
      .trim();

    const finalTitle =
      (cleanTitle || clause).charAt(0).toUpperCase() +
      (cleanTitle || clause).slice(1);

    return {
      voice_note_id: voiceNoteId || null,
      action_type: 'planner_task',
      title: finalTitle,
      description: `Voice Note: "${clause}"`,
      due_date: dueDate,
      priority,
      plan_id: settings.defaultPlannerPlanId || undefined,
      plan_name: settings.defaultPlannerPlanName || undefined,
      bucket_id: settings.defaultPlannerBucketId || undefined,
      bucket_name: settings.defaultPlannerBucketName || undefined,
      status: 'draft',
    };
  }

  // 5. Parse Outlook Calendar Event
  if (actionType === 'outlook_event') {
    const cleanTitle = clause
      .replace(
        /^(?:please\s+)?(?:can\s+you\s+)?(?:schedule|book|create|add|set\s+up)\s+(?:a\s+|an\s+)?(?:new\s+)?(?:outlook\s+)?(?:calendar\s+event|meeting|appointment|call|event)\s*(?:with\s+|for\s+|to\s+|about\s+)?/i,
        ''
      )
      .trim();

    const matchedContact = findMatchingContact(undefined, clause, contacts);

    return {
      voice_note_id: voiceNoteId || null,
      action_type: 'outlook_event',
      title:
        (cleanTitle || clause).charAt(0).toUpperCase() +
        (cleanTitle || clause).slice(1),
      description: `Scheduled from voice note: "${clause}"`,
      recipient_name: matchedContact?.name,
      recipient_email: matchedContact?.email,
      due_date: dueDate || formatDateISO(new Date()),
      start_time: startTime || '09:00',
      priority,
      status: 'draft',
    };
  }

  // 6. Default: Outlook Task / Microsoft To Do
  const cleanTaskTitle = clause
    .replace(
      /^(?:please\s+)?(?:can\s+you\s+)?(?:create|add|generate|make|set|remind\s+me\s+to)\s+(?:a\s+|an\s+)?(?:new\s+)?(?:outlook\s+task|task\s+in\s+outlook|to\s*do\s+task|task|reminder)\s*(?:in\s+outlook\s*)?(?:to\s+|for\s+|that\s+)?/i,
      ''
    )
    .replace(/\bin\s+outlook\b/gi, '')
    .trim();

  const finalTaskTitle =
    (cleanTaskTitle || clause).charAt(0).toUpperCase() +
    (cleanTaskTitle || clause).slice(1);

  return {
    voice_note_id: voiceNoteId || null,
    action_type: 'outlook_task',
    title: finalTaskTitle,
    description: `Voice Note: "${clause}"`,
    due_date: dueDate,
    priority,
    status: 'draft',
  };
}

/**
 * Parses a voice transcript using AI (if configured) or the instant local multi-intent parser.
 */
export async function parseVoiceTranscriptToActions(
  transcript: string,
  contacts: Contact[],
  settings: AssistantSettings,
  voiceNoteId?: string
): Promise<DraftActionInput[]> {
  const cleanTranscript = transcript.trim();
  if (!cleanTranscript) return [];

  // Try AI provider if user selected Gemini or OpenAI (or provided a key)
  if (
    (settings.aiProvider === 'gemini' && settings.geminiApiKey) ||
    (settings.aiProvider === 'openai' && settings.openaiApiKey)
  ) {
    try {
      const response = await fetch('/api/ai/parse', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          transcript: cleanTranscript,
          contacts: contacts.map((c) => ({
            name: c.name,
            phone: c.phone,
            email: c.email,
          })),
          provider: settings.aiProvider,
          apiKey:
            settings.aiProvider === 'gemini'
              ? settings.geminiApiKey
              : settings.openaiApiKey,
          todayISO: new Date().toISOString().split('T')[0],
        }),
      });

      if (response.ok) {
        const data = await response.json();
        if (Array.isArray(data.actions) && data.actions.length > 0) {
          return data.actions.map((item: any) => {
            const matched = findMatchingContact(
              item.recipient_name,
              cleanTranscript,
              contacts
            );
            return {
              voice_note_id: voiceNoteId || null,
              action_type: (item.action_type as ActionType) || 'outlook_task',
              title: item.title || cleanTranscript,
              description: item.description || cleanTranscript,
              recipient_name: item.recipient_name || matched?.name || '',
              recipient_phone: item.recipient_phone || matched?.phone || '',
              recipient_email: item.recipient_email || matched?.email || '',
              due_date: item.due_date || undefined,
              start_time: item.start_time || undefined,
              priority: (item.priority as ActionPriority) || 'normal',
              plan_id: settings.defaultPlannerPlanId || undefined,
              plan_name: settings.defaultPlannerPlanName || undefined,
              bucket_id: settings.defaultPlannerBucketId || undefined,
              bucket_name: settings.defaultPlannerBucketName || undefined,
              status: 'draft',
            };
          });
        }
      }
    } catch (err) {
      console.warn('AI parse fallback to smart local parser:', err);
    }
  }

  // Smart Local Multi-Action Parser
  const segments = splitIntoSegments(cleanTranscript);
  return segments.map((seg) =>
    parseSingleClause(seg, contacts, settings, voiceNoteId)
  );
}
