export type ActionType =
  | 'whatsapp'
  | 'outlook_task'
  | 'planner_task'
  | 'outlook_email'
  | 'outlook_event';

export type ActionStatus = 'draft' | 'queued' | 'completed' | 'failed';
export type ActionPriority = 'low' | 'normal' | 'high' | 'urgent';

export interface Contact {
  id: string;
  name: string;
  phone?: string;
  email?: string;
  role?: string;
  company?: string;
  notes?: string;
  created_at?: string;
  updated_at?: string;
}

export interface VoiceNote {
  id: string;
  transcript: string;
  summary?: string;
  duration_seconds: number;
  language: string;
  ai_provider: string;
  actions_count: number;
  created_at: string;
}

export interface AssistantAction {
  id: string;
  voice_note_id?: string | null;
  action_type: ActionType;
  title: string;
  description?: string;
  recipient_name?: string;
  recipient_phone?: string;
  recipient_email?: string;
  due_date?: string;
  start_time?: string;
  end_time?: string;
  priority: ActionPriority;
  plan_id?: string;
  plan_name?: string;
  bucket_id?: string;
  bucket_name?: string;
  status: ActionStatus;
  external_id?: string;
  external_url?: string;
  error_message?: string;
  executed_at?: string;
  created_at: string;
  updated_at?: string;
}

export interface PlannerPlan {
  id: string;
  title: string;
  owner?: string;
}

export interface PlannerBucket {
  id: string;
  name: string;
  planId: string;
}

export interface TodoList {
  id: string;
  displayName: string;
  wellknownListName?: string;
}

export interface AssistantSettings {
  // WhatsApp Background API Settings
  whatsappProvider: 'meta' | 'twilio';
  metaAccessToken: string;
  metaPhoneNumberId: string;
  metaTemplateName?: string; // Optional template fallback
  twilioAccountSid: string;
  twilioAuthToken: string;
  twilioFromNumber: string; // e.g. whatsapp:+14155238886
  defaultCountryCode: string; // e.g. +27

  // Microsoft 365 Settings
  azureClientId: string;
  azureTenantId: string;
  defaultPlannerPlanId: string;
  defaultPlannerPlanName: string;
  defaultPlannerBucketId: string;
  defaultPlannerBucketName: string;
  defaultTodoListId: string;
  emailMode: 'send_immediately' | 'create_draft';

  // AI & Voice Settings
  speechLanguage: string; // e.g. 'en-ZA', 'en-US', 'af-ZA'
  aiProvider: 'smart-local' | 'gemini' | 'openai';
  geminiApiKey: string;
  openaiApiKey: string;
  autoExecuteHighConfidence: boolean;
}

export const DEFAULT_SETTINGS: AssistantSettings = {
  whatsappProvider: 'meta',
  metaAccessToken: '',
  metaPhoneNumberId: '',
  metaTemplateName: '',
  twilioAccountSid: '',
  twilioAuthToken: '',
  twilioFromNumber: 'whatsapp:+14155238886',
  defaultCountryCode: '+27',

  azureClientId:
    process.env.NEXT_PUBLIC_AZURE_CLIENT_ID || 'abae4ffc-842f-45cb-aea5-8394d2de9a82',
  azureTenantId:
    process.env.NEXT_PUBLIC_AZURE_TENANT_ID || '70fc3e15-9b2c-4e05-8f43-9cebcc9bfada',
  defaultPlannerPlanId: '',
  defaultPlannerPlanName: '',
  defaultPlannerBucketId: '',
  defaultPlannerBucketName: '',
  defaultTodoListId: '',
  emailMode: 'send_immediately',

  speechLanguage: 'en-ZA',
  aiProvider: 'smart-local',
  geminiApiKey: '',
  openaiApiKey: '',
  autoExecuteHighConfidence: false,
};
