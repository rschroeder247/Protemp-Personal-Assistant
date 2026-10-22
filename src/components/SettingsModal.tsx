'use client';

import React, { useState, useEffect } from 'react';
import {
  AssistantSettings,
  PlannerBucket,
  PlannerPlan,
  TodoList,
} from '@/lib/types';
import {
  Settings,
  X,
  MessageSquare,
  LayoutGrid,
  Sparkles,
  Save,
  CheckCircle2,
} from 'lucide-react';

interface SettingsModalProps {
  isOpen: boolean;
  onClose: () => void;
  settings: AssistantSettings;
  onSave: (next: AssistantSettings) => Promise<void>;
  plannerPlans: PlannerPlan[];
  todoLists: TodoList[];
  bucketsByPlan: Record<string, PlannerBucket[]>;
  onLoadBuckets: (planId: string) => Promise<void>;
}

export default function SettingsModal({
  isOpen,
  onClose,
  settings,
  onSave,
  plannerPlans,
  todoLists,
  bucketsByPlan,
  onLoadBuckets,
}: SettingsModalProps) {
  const [form, setForm] = useState<AssistantSettings>(settings);
  const [activeTab, setActiveTab] = useState<'whatsapp' | 'm365' | 'ai'>('whatsapp');
  const [savedNotice, setSavedNotice] = useState(false);

  useEffect(() => {
    setForm(settings);
  }, [settings, isOpen]);

  useEffect(() => {
    if (form.defaultPlannerPlanId) {
      onLoadBuckets(form.defaultPlannerPlanId);
    }
  }, [form.defaultPlannerPlanId]);

  if (!isOpen) return null;

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    await onSave(form);
    setSavedNotice(true);
    setTimeout(() => {
      setSavedNotice(false);
      onClose();
    }, 800);
  };

  const currentBuckets = form.defaultPlannerPlanId
    ? bucketsByPlan[form.defaultPlannerPlanId] || []
    : [];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 backdrop-blur-sm p-4">
      <div className="flex max-h-[90vh] w-full max-w-2xl flex-col rounded-2xl border border-slate-800 bg-slate-900 shadow-2xl">
        {/* Header */}
        <div className="flex items-center justify-between border-b border-slate-800 px-6 py-4">
          <div className="flex items-center gap-2.5">
            <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-indigo-500/20 text-indigo-400">
              <Settings className="h-5 w-5" />
            </div>
            <div>
              <h3 className="text-base font-semibold text-white">
                Assistant Integrations &amp; API Settings
              </h3>
              <p className="text-xs text-slate-400">
                Configure WhatsApp Business Cloud / Twilio API, Microsoft 365 defaults, and AI engines.
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="rounded-lg p-1.5 text-slate-400 hover:bg-slate-800 hover:text-white"
          >
            <X className="h-5 w-5" />
          </button>
        </div>

        {/* Tabs */}
        <div className="flex border-b border-slate-800 px-6 bg-slate-950/40">
          <button
            type="button"
            onClick={() => setActiveTab('whatsapp')}
            className={`flex items-center gap-2 border-b-2 px-4 py-3 text-xs font-semibold transition ${
              activeTab === 'whatsapp'
                ? 'border-emerald-500 text-emerald-400'
                : 'border-transparent text-slate-400 hover:text-slate-200'
            }`}
          >
            <MessageSquare className="h-4 w-4" />
            WhatsApp Cloud / Twilio API
          </button>
          <button
            type="button"
            onClick={() => setActiveTab('m365')}
            className={`flex items-center gap-2 border-b-2 px-4 py-3 text-xs font-semibold transition ${
              activeTab === 'm365'
                ? 'border-sky-500 text-sky-400'
                : 'border-transparent text-slate-400 hover:text-slate-200'
            }`}
          >
            <LayoutGrid className="h-4 w-4" />
            Microsoft 365 (Outlook &amp; Planner)
          </button>
          <button
            type="button"
            onClick={() => setActiveTab('ai')}
            className={`flex items-center gap-2 border-b-2 px-4 py-3 text-xs font-semibold transition ${
              activeTab === 'ai'
                ? 'border-indigo-500 text-indigo-400'
                : 'border-transparent text-slate-400 hover:text-slate-200'
            }`}
          >
            <Sparkles className="h-4 w-4" />
            Voice &amp; AI Parser
          </button>
        </div>

        {/* Form Body */}
        <form onSubmit={handleSubmit} className="flex-1 overflow-y-auto p-6 space-y-5">
          {activeTab === 'whatsapp' && (
            <div className="space-y-4">
              <div>
                <label className="block text-xs font-medium text-slate-300 mb-2">
                  Background WhatsApp Provider
                </label>
                <div className="grid grid-cols-2 gap-3">
                  <button
                    type="button"
                    onClick={() =>
                      setForm({ ...form, whatsappProvider: 'meta' })
                    }
                    className={`rounded-xl border p-3 text-left transition ${
                      form.whatsappProvider === 'meta'
                        ? 'border-emerald-500 bg-emerald-500/10 text-white'
                        : 'border-slate-800 bg-slate-950/60 text-slate-400'
                    }`}
                  >
                    <div className="text-xs font-semibold">
                      Meta WhatsApp Business Cloud API
                    </div>
                    <div className="text-[11px] opacity-80 mt-0.5">
                      Official Meta Graph API (graph.facebook.com)
                    </div>
                  </button>

                  <button
                    type="button"
                    onClick={() =>
                      setForm({ ...form, whatsappProvider: 'twilio' })
                    }
                    className={`rounded-xl border p-3 text-left transition ${
                      form.whatsappProvider === 'twilio'
                        ? 'border-emerald-500 bg-emerald-500/10 text-white'
                        : 'border-slate-800 bg-slate-950/60 text-slate-400'
                    }`}
                  >
                    <div className="text-xs font-semibold">
                      Twilio WhatsApp Messaging API
                    </div>
                    <div className="text-[11px] opacity-80 mt-0.5">
                      Twilio Programmable Messaging (api.twilio.com)
                    </div>
                  </button>
                </div>
              </div>

              {form.whatsappProvider === 'meta' ? (
                <div className="space-y-3 rounded-xl border border-slate-800 bg-slate-950/60 p-4">
                  <div>
                    <label className="block text-xs font-medium text-slate-300 mb-1">
                      Meta Phone Number ID
                    </label>
                    <input
                      type="text"
                      value={form.metaPhoneNumberId}
                      onChange={(e) =>
                        setForm({ ...form, metaPhoneNumberId: e.target.value })
                      }
                      placeholder="e.g. 105928374918273"
                      className="w-full rounded-lg border border-slate-800 bg-slate-900 px-3 py-2 text-xs text-white focus:border-emerald-500 focus:outline-none"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-medium text-slate-300 mb-1">
                      Meta Permanent / System User Access Token
                    </label>
                    <input
                      type="password"
                      value={form.metaAccessToken}
                      onChange={(e) =>
                        setForm({ ...form, metaAccessToken: e.target.value })
                      }
                      placeholder="EAAG..."
                      className="w-full rounded-lg border border-slate-800 bg-slate-900 px-3 py-2 text-xs text-white focus:border-emerald-500 focus:outline-none"
                    />
                  </div>
                </div>
              ) : (
                <div className="space-y-3 rounded-xl border border-slate-800 bg-slate-950/60 p-4">
                  <div>
                    <label className="block text-xs font-medium text-slate-300 mb-1">
                      Twilio Account SID
                    </label>
                    <input
                      type="text"
                      value={form.twilioAccountSid}
                      onChange={(e) =>
                        setForm({ ...form, twilioAccountSid: e.target.value })
                      }
                      placeholder="ACxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
                      className="w-full rounded-lg border border-slate-800 bg-slate-900 px-3 py-2 text-xs text-white focus:border-emerald-500 focus:outline-none"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-medium text-slate-300 mb-1">
                      Twilio Auth Token
                    </label>
                    <input
                      type="password"
                      value={form.twilioAuthToken}
                      onChange={(e) =>
                        setForm({ ...form, twilioAuthToken: e.target.value })
                      }
                      placeholder="Your Twilio Auth Token"
                      className="w-full rounded-lg border border-slate-800 bg-slate-900 px-3 py-2 text-xs text-white focus:border-emerald-500 focus:outline-none"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-medium text-slate-300 mb-1">
                      Twilio WhatsApp Sender Number
                    </label>
                    <input
                      type="text"
                      value={form.twilioFromNumber}
                      onChange={(e) =>
                        setForm({ ...form, twilioFromNumber: e.target.value })
                      }
                      placeholder="whatsapp:+14155238886"
                      className="w-full rounded-lg border border-slate-800 bg-slate-900 px-3 py-2 text-xs text-white focus:border-emerald-500 focus:outline-none"
                    />
                  </div>
                </div>
              )}

              <div>
                <label className="block text-xs font-medium text-slate-300 mb-1">
                  Default Country Dialing Code (for local numbers like 082...)
                </label>
                <input
                  type="text"
                  value={form.defaultCountryCode}
                  onChange={(e) =>
                    setForm({ ...form, defaultCountryCode: e.target.value })
                  }
                  placeholder="+27"
                  className="w-40 rounded-lg border border-slate-800 bg-slate-900 px-3 py-2 text-xs text-white focus:border-emerald-500 focus:outline-none"
                />
              </div>
            </div>
          )}

          {activeTab === 'm365' && (
            <div className="space-y-4">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-medium text-slate-300 mb-1">
                    Azure AD App Client ID
                  </label>
                  <input
                    type="text"
                    value={form.azureClientId}
                    onChange={(e) =>
                      setForm({ ...form, azureClientId: e.target.value })
                    }
                    className="w-full rounded-lg border border-slate-800 bg-slate-950 px-3 py-2 text-xs text-white focus:border-sky-500 focus:outline-none"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-300 mb-1">
                    Azure AD Tenant ID (Protemp)
                  </label>
                  <input
                    type="text"
                    value={form.azureTenantId}
                    onChange={(e) =>
                      setForm({ ...form, azureTenantId: e.target.value })
                    }
                    className="w-full rounded-lg border border-slate-800 bg-slate-950 px-3 py-2 text-xs text-white focus:border-sky-500 focus:outline-none"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-medium text-slate-300 mb-1">
                    Default Microsoft Planner Plan
                  </label>
                  <select
                    value={form.defaultPlannerPlanId}
                    onChange={(e) => {
                      const pid = e.target.value;
                      const pObj = plannerPlans.find((p) => p.id === pid);
                      setForm({
                        ...form,
                        defaultPlannerPlanId: pid,
                        defaultPlannerPlanName: pObj?.title || '',
                        defaultPlannerBucketId: '',
                      });
                    }}
                    className="w-full rounded-lg border border-slate-800 bg-slate-950 px-3 py-2 text-xs text-white focus:border-sky-500 focus:outline-none"
                  >
                    <option value="">-- First Available Plan --</option>
                    {plannerPlans.map((p) => (
                      <option key={p.id} value={p.id}>
                        {p.title}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-medium text-slate-300 mb-1">
                    Default Planner Bucket
                  </label>
                  <select
                    value={form.defaultPlannerBucketId}
                    onChange={(e) => {
                      const bid = e.target.value;
                      const bObj = currentBuckets.find((b) => b.id === bid);
                      setForm({
                        ...form,
                        defaultPlannerBucketId: bid,
                        defaultPlannerBucketName: bObj?.name || '',
                      });
                    }}
                    className="w-full rounded-lg border border-slate-800 bg-slate-950 px-3 py-2 text-xs text-white focus:border-sky-500 focus:outline-none"
                  >
                    <option value="">-- Default First Bucket --</option>
                    {currentBuckets.map((b) => (
                      <option key={b.id} value={b.id}>
                        {b.name}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-medium text-slate-300 mb-1">
                    Default Outlook / To Do List
                  </label>
                  <select
                    value={form.defaultTodoListId}
                    onChange={(e) =>
                      setForm({ ...form, defaultTodoListId: e.target.value })
                    }
                    className="w-full rounded-lg border border-slate-800 bg-slate-950 px-3 py-2 text-xs text-white focus:border-sky-500 focus:outline-none"
                  >
                    <option value="">-- Default &quot;Tasks&quot; List --</option>
                    {todoLists.map((l) => (
                      <option key={l.id} value={l.id}>
                        {l.displayName}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-medium text-slate-300 mb-1">
                    Outlook Email Behavior
                  </label>
                  <select
                    value={form.emailMode}
                    onChange={(e) =>
                      setForm({
                        ...form,
                        emailMode: e.target.value as
                          | 'send_immediately'
                          | 'create_draft',
                      })
                    }
                    className="w-full rounded-lg border border-slate-800 bg-slate-950 px-3 py-2 text-xs text-white focus:border-sky-500 focus:outline-none"
                  >
                    <option value="send_immediately">
                      Send Immediately via Graph / Outlook
                    </option>
                    <option value="create_draft">
                      Save to Outlook Drafts First
                    </option>
                  </select>
                </div>
              </div>
            </div>
          )}

          {activeTab === 'ai' && (
            <div className="space-y-4">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-medium text-slate-300 mb-1">
                    Speech Recognition Language
                  </label>
                  <select
                    value={form.speechLanguage}
                    onChange={(e) =>
                      setForm({ ...form, speechLanguage: e.target.value })
                    }
                    className="w-full rounded-lg border border-slate-800 bg-slate-950 px-3 py-2 text-xs text-white focus:border-indigo-500 focus:outline-none"
                  >
                    <option value="en-ZA">English (South Africa - en-ZA)</option>
                    <option value="en-GB">English (UK - en-GB)</option>
                    <option value="en-US">English (US - en-US)</option>
                    <option value="af-ZA">Afrikaans (South Africa - af-ZA)</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-medium text-slate-300 mb-1">
                    Intent Extraction Engine
                  </label>
                  <select
                    value={form.aiProvider}
                    onChange={(e) =>
                      setForm({
                        ...form,
                        aiProvider: e.target.value as
                          | 'smart-local'
                          | 'gemini'
                          | 'openai',
                      })
                    }
                    className="w-full rounded-lg border border-slate-800 bg-slate-950 px-3 py-2 text-xs text-white focus:border-indigo-500 focus:outline-none"
                  >
                    <option value="smart-local">
                      Built-in Smart Multi-Action Parser (Free, Instant)
                    </option>
                    <option value="gemini">
                      Google Gemini 2.0 Flash + Smart Fallback
                    </option>
                    <option value="openai">
                      OpenAI GPT-4o-mini + Smart Fallback
                    </option>
                  </select>
                </div>
              </div>

              {form.aiProvider === 'gemini' && (
                <div>
                  <label className="block text-xs font-medium text-slate-300 mb-1">
                    Google Gemini API Key
                  </label>
                  <input
                    type="password"
                    value={form.geminiApiKey}
                    onChange={(e) =>
                      setForm({ ...form, geminiApiKey: e.target.value })
                    }
                    placeholder="AIzaSy..."
                    className="w-full rounded-lg border border-slate-800 bg-slate-950 px-3 py-2 text-xs text-white focus:border-indigo-500 focus:outline-none"
                  />
                </div>
              )}

              {form.aiProvider === 'openai' && (
                <div>
                  <label className="block text-xs font-medium text-slate-300 mb-1">
                    OpenAI API Key
                  </label>
                  <input
                    type="password"
                    value={form.openaiApiKey}
                    onChange={(e) =>
                      setForm({ ...form, openaiApiKey: e.target.value })
                    }
                    placeholder="sk-..."
                    className="w-full rounded-lg border border-slate-800 bg-slate-950 px-3 py-2 text-xs text-white focus:border-indigo-500 focus:outline-none"
                  />
                </div>
              )}

              <label className="flex items-center gap-2.5 pt-2 cursor-pointer">
                <input
                  type="checkbox"
                  checked={form.autoExecuteHighConfidence}
                  onChange={(e) =>
                    setForm({
                      ...form,
                      autoExecuteHighConfidence: e.target.checked,
                    })
                  }
                  className="h-4 w-4 rounded border-slate-700 bg-slate-900 text-indigo-600 focus:ring-indigo-500"
                />
                <span className="text-xs text-slate-300">
                  Automatically execute extracted actions immediately after voice note finishes (Zero-Click Mode)
                </span>
              </label>
            </div>
          )}

          {/* Footer */}
          <div className="flex items-center justify-end gap-3 pt-4 border-t border-slate-800">
            <button
              type="button"
              onClick={onClose}
              className="rounded-xl border border-slate-700 bg-slate-800 hover:bg-slate-700 px-4 py-2 text-xs font-medium text-slate-300"
            >
              Cancel
            </button>
            <button
              type="submit"
              className="inline-flex items-center gap-1.5 rounded-xl bg-indigo-600 hover:bg-indigo-500 px-5 py-2 text-xs font-semibold text-white shadow-md"
            >
              {savedNotice ? (
                <>
                  <CheckCircle2 className="h-4 w-4 text-emerald-300" /> Saved!
                </>
              ) : (
                <>
                  <Save className="h-4 w-4" /> Save Settings
                </>
              )}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
