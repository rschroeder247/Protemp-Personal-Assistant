-- Personal Assistant Schema for Protemp
CREATE TABLE IF NOT EXISTS public.pa_contacts (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  name TEXT NOT NULL,
  phone TEXT,
  email TEXT,
  role TEXT,
  company TEXT DEFAULT 'Protemp',
  notes TEXT,
  created_at TIMESTAMPTZ DEFAULT now(),
  updated_at TIMESTAMPTZ DEFAULT now()
);

CREATE TABLE IF NOT EXISTS public.pa_voice_notes (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  transcript TEXT NOT NULL,
  summary TEXT,
  duration_seconds INTEGER DEFAULT 0,
  language TEXT DEFAULT 'en-ZA',
  ai_provider TEXT DEFAULT 'smart-parser',
  actions_count INTEGER DEFAULT 0,
  created_at TIMESTAMPTZ DEFAULT now()
);

CREATE TABLE IF NOT EXISTS public.pa_actions (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  voice_note_id UUID REFERENCES public.pa_voice_notes(id) ON DELETE SET NULL,
  action_type TEXT NOT NULL CHECK (action_type IN ('whatsapp', 'outlook_task', 'planner_task', 'outlook_email', 'outlook_event')),
  title TEXT NOT NULL,
  description TEXT,
  recipient_name TEXT,
  recipient_phone TEXT,
  recipient_email TEXT,
  due_date TEXT,
  start_time TEXT,
  end_time TEXT,
  priority TEXT DEFAULT 'normal',
  plan_id TEXT,
  plan_name TEXT,
  bucket_id TEXT,
  bucket_name TEXT,
  status TEXT NOT NULL DEFAULT 'draft' CHECK (status IN ('draft', 'queued', 'completed', 'failed')),
  external_id TEXT,
  external_url TEXT,
  error_message TEXT,
  executed_at TIMESTAMPTZ,
  created_at TIMESTAMPTZ DEFAULT now(),
  updated_at TIMESTAMPTZ DEFAULT now()
);

CREATE TABLE IF NOT EXISTS public.pa_settings (
  key TEXT PRIMARY KEY,
  value JSONB NOT NULL DEFAULT '{}'::jsonb,
  updated_at TIMESTAMPTZ DEFAULT now()
);

ALTER TABLE public.pa_contacts ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.pa_voice_notes ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.pa_actions ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.pa_settings ENABLE ROW LEVEL SECURITY;
