import { createClient } from '@supabase/supabase-js';
import {
  AssistantAction,
  AssistantSettings,
  Contact,
  DEFAULT_SETTINGS,
  VoiceNote,
} from './types';

const supabaseUrl =
  process.env.NEXT_PUBLIC_SUPABASE_URL ||
  'https://dphxaglpramhwcwcnpay.supabase.co';

const supabaseAnonKey =
  process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY ||
  'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRwaHhhZ2xwcmFtaHdjd2NucGF5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzMyNDg4NzksImV4cCI6MjA4ODgyNDg3OX0.Z2uLJlqYfs6RGcrRe3SpyaHo7BzvO4QgmYp9lXVEKN0';

export const supabase = createClient(supabaseUrl, supabaseAnonKey);

const LOCAL_SETTINGS_KEY = 'pa_assistant_settings_v1';

// ==========================================
// CONTACTS
// ==========================================
export async function fetchContacts(): Promise<Contact[]> {
  const { data, error } = await supabase
    .from('pa_contacts')
    .select('*')
    .order('name', { ascending: true });

  if (error) {
    console.warn('Supabase fetchContacts error:', error.message);
    return [];
  }
  return (data || []) as Contact[];
}

export async function upsertContact(
  contact: Partial<Contact> & { name: string }
): Promise<Contact | null> {
  const payload = {
    ...contact,
    updated_at: new Date().toISOString(),
  };

  const { data, error } = await supabase
    .from('pa_contacts')
    .upsert(payload)
    .select()
    .single();

  if (error) {
    console.error('Supabase upsertContact error:', error.message);
    throw new Error(error.message);
  }
  return data as Contact;
}

export async function deleteContactById(id: string): Promise<void> {
  const { error } = await supabase.from('pa_contacts').delete().eq('id', id);
  if (error) {
    throw new Error(error.message);
  }
}

// ==========================================
// VOICE NOTES
// ==========================================
export async function fetchVoiceNotes(): Promise<VoiceNote[]> {
  const { data, error } = await supabase
    .from('pa_voice_notes')
    .select('*')
    .order('created_at', { ascending: false })
    .limit(50);

  if (error) {
    console.warn('Supabase fetchVoiceNotes error:', error.message);
    return [];
  }
  return (data || []) as VoiceNote[];
}

export async function createVoiceNote(
  note: Omit<VoiceNote, 'id' | 'created_at'>
): Promise<VoiceNote | null> {
  const { data, error } = await supabase
    .from('pa_voice_notes')
    .insert(note)
    .select()
    .single();

  if (error) {
    console.error('Supabase createVoiceNote error:', error.message);
    return null;
  }
  return data as VoiceNote;
}

export async function deleteVoiceNoteById(id: string): Promise<void> {
  const { error } = await supabase.from('pa_voice_notes').delete().eq('id', id);
  if (error) throw new Error(error.message);
}

// ==========================================
// ACTIONS (WHATSAPP, OUTLOOK, PLANNER, EMAIL)
// ==========================================
export async function fetchActions(): Promise<AssistantAction[]> {
  const { data, error } = await supabase
    .from('pa_actions')
    .select('*')
    .order('created_at', { ascending: false })
    .limit(100);

  if (error) {
    console.warn('Supabase fetchActions error:', error.message);
    return [];
  }
  return (data || []) as AssistantAction[];
}

export async function createActionsBatch(
  actions: Omit<AssistantAction, 'id' | 'created_at' | 'updated_at'>[]
): Promise<AssistantAction[]> {
  if (actions.length === 0) return [];
  const { data, error } = await supabase
    .from('pa_actions')
    .insert(actions)
    .select();

  if (error) {
    console.error('Supabase createActionsBatch error:', error.message);
    throw new Error(error.message);
  }
  return (data || []) as AssistantAction[];
}

export async function updateActionRecord(
  id: string,
  patch: Partial<AssistantAction>
): Promise<AssistantAction> {
  const { data, error } = await supabase
    .from('pa_actions')
    .update({
      ...patch,
      updated_at: new Date().toISOString(),
    })
    .eq('id', id)
    .select()
    .single();

  if (error) {
    console.error('Supabase updateActionRecord error:', error.message);
    throw new Error(error.message);
  }
  return data as AssistantAction;
}

export async function deleteActionRecord(id: string): Promise<void> {
  const { error } = await supabase.from('pa_actions').delete().eq('id', id);
  if (error) {
    throw new Error(error.message);
  }
}

// ==========================================
// SETTINGS (SUPABASE + LOCALSTORAGE SYNC)
// ==========================================
export async function loadAssistantSettings(): Promise<AssistantSettings> {
  let local: Partial<AssistantSettings> = {};
  if (typeof window !== 'undefined') {
    try {
      const raw = localStorage.getItem(LOCAL_SETTINGS_KEY);
      if (raw) local = JSON.parse(raw);
    } catch {
      // ignore
    }
  }

  try {
    const { data, error } = await supabase
      .from('pa_settings')
      .select('value')
      .eq('key', 'default_profile')
      .maybeSingle();

    if (!error && data?.value) {
      const merged: AssistantSettings = {
        ...DEFAULT_SETTINGS,
        ...(data.value as Partial<AssistantSettings>),
        ...local,
      };
      if (typeof window !== 'undefined') {
        localStorage.setItem(LOCAL_SETTINGS_KEY, JSON.stringify(merged));
      }
      return merged;
    }
  } catch {
    // ignore
  }

  return { ...DEFAULT_SETTINGS, ...local };
}

export async function saveAssistantSettings(
  settings: AssistantSettings
): Promise<void> {
  if (typeof window !== 'undefined') {
    try {
      localStorage.setItem(LOCAL_SETTINGS_KEY, JSON.stringify(settings));
    } catch {
      // ignore
    }
  }

  try {
    await supabase.from('pa_settings').upsert({
      key: 'default_profile',
      value: settings,
      updated_at: new Date().toISOString(),
    });
  } catch (err) {
    console.warn('Could not sync settings to Supabase:', err);
  }
}
