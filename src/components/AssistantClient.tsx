'use client';

import React, { useState, useEffect, useCallback } from 'react';
import {
  ActionType,
  AssistantAction,
  AssistantSettings,
  Contact,
  DEFAULT_SETTINGS,
  PlannerBucket,
  PlannerPlan,
  TodoList,
  VoiceNote,
} from '@/lib/types';
import {
  createActionsBatch,
  createVoiceNote,
  deleteActionRecord,
  deleteContactById,
  deleteVoiceNoteById,
  fetchActions,
  fetchContacts,
  fetchVoiceNotes,
  loadAssistantSettings,
  saveAssistantSettings,
  updateActionRecord,
  upsertContact,
} from '@/lib/supabase';
import { graphService } from '@/lib/graphService';
import { setStoredAzureConfig } from '@/lib/msalConfig';
import { parseVoiceTranscriptToActions } from '@/lib/voiceParser';
import VoiceRecorderCard from './VoiceRecorderCard';
import ActionCard from './ActionCard';
import ContactsModal from './ContactsModal';
import SettingsModal from './SettingsModal';
import {
  Sparkles,
  Users,
  Settings,
  LogIn,
  LogOut,
  MessageSquare,
  CheckSquare,
  LayoutGrid,
  Mail,
  Plus,
  Play,
  RefreshCw,
  FileAudio,
  CheckCircle2,
  Clock,
  ExternalLink,
  Trash2,
} from 'lucide-react';

export default function AssistantClient() {
  const [actions, setActions] = useState<AssistantAction[]>([]);
  const [voiceNotes, setVoiceNotes] = useState<VoiceNote[]>([]);
  const [contacts, setContacts] = useState<Contact[]>([]);
  const [settings, setSettings] = useState<AssistantSettings>(DEFAULT_SETTINGS);

  // Microsoft 365 State
  const [m365User, setM365User] = useState<{
    name: string;
    email: string;
  } | null>(null);
  const [plannerPlans, setPlannerPlans] = useState<PlannerPlan[]>([]);
  const [todoLists, setTodoLists] = useState<TodoList[]>([]);
  const [bucketsByPlan, setBucketsByPlan] = useState<
    Record<string, PlannerBucket[]>
  >({});
  const [liveTodoTasks, setLiveTodoTasks] = useState<any[]>([]);
  const [livePlannerTasks, setLivePlannerTasks] = useState<any[]>([]);
  const [isLoadingLiveM365, setIsLoadingLiveM365] = useState(false);

  // UI State
  const [activeWorkspaceTab, setActiveWorkspaceTab] = useState<
    'actions' | 'live_m365' | 'transcripts'
  >('actions');
  const [actionFilter, setActionFilter] = useState<
    'all' | 'draft' | 'whatsapp' | 'outlook_task' | 'planner_task' | 'outlook_email' | 'completed'
  >('all');
  const [isProcessingVoice, setIsProcessingVoice] = useState(false);
  const [executingIds, setExecutingIds] = useState<Record<string, boolean>>({});
  const [isContactsOpen, setIsContactsOpen] = useState(false);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isM365ModalOpen, setIsM365ModalOpen] = useState(false);
  const [toast, setToast] = useState<{
    type: 'success' | 'error' | 'info';
    text: string;
  } | null>(null);

  const showToast = (
    text: string,
    type: 'success' | 'error' | 'info' = 'info'
  ) => {
    setToast({ text, type });
    setTimeout(() => {
      setToast((prev) => (prev?.text === text ? null : prev));
    }, 5000);
  };

  // Load initial data from Supabase & check M365 login
  const loadInitialData = useCallback(async () => {
    const [loadedSettings, loadedContacts, loadedActions, loadedNotes] =
      await Promise.all([
        loadAssistantSettings(),
        fetchContacts(),
        fetchActions(),
        fetchVoiceNotes(),
      ]);

    setSettings(loadedSettings);
    setContacts(loadedContacts);
    setActions(loadedActions);
    setVoiceNotes(loadedNotes);

    // Check Microsoft 365 MSAL account
    try {
      const account = await graphService.getAccount();
      if (account) {
        setM365User({
          name: account.name || account.username,
          email: account.username,
        });
        loadM365Metadata();
      }
    } catch (err) {
      console.warn('MSAL init check:', err);
    }
  }, []);

  useEffect(() => {
    loadInitialData();
  }, [loadInitialData]);

  const loadM365Metadata = async () => {
    try {
      const [plans, lists] = await Promise.all([
        graphService.getPlannerPlans().catch(() => []),
        graphService.getTodoLists().catch(() => []),
      ]);
      setPlannerPlans(plans);
      setTodoLists(lists);
    } catch (err) {
      console.warn('Could not load M365 metadata:', err);
    }
  };

  const handleLoadBuckets = async (planId: string) => {
    if (!planId || bucketsByPlan[planId]) return;
    const buckets = await graphService.getPlannerBuckets(planId);
    setBucketsByPlan((prev) => ({ ...prev, [planId]: buckets }));
  };

  const refreshLiveM365Tasks = async () => {
    if (!m365User) return;
    setIsLoadingLiveM365(true);
    try {
      const [todoItems, plannerItems] = await Promise.all([
        graphService.getLiveOutlookTasks().catch(() => []),
        graphService.getLivePlannerTasks().catch(() => []),
      ]);
      setLiveTodoTasks(todoItems);
      setLivePlannerTasks(plannerItems);
    } finally {
      setIsLoadingLiveM365(false);
    }
  };

  useEffect(() => {
    if (activeWorkspaceTab === 'live_m365' && m365User) {
      refreshLiveM365Tasks();
    }
  }, [activeWorkspaceTab, m365User]);

  // ==========================================================
  // EXECUTE SINGLE ACTION (WHATSAPP, OUTLOOK, PLANNER, EMAIL)
  // ==========================================================
  const handleExecuteAction = async (action: AssistantAction) => {
    setExecutingIds((prev) => ({ ...prev, [action.id]: true }));
    try {
      let externalId = '';
      let externalUrl = '';

      if (action.action_type === 'whatsapp') {
        const phone = action.recipient_phone?.trim();
        if (!phone) {
          throw new Error(
            `Please enter a WhatsApp phone number for "${
              action.recipient_name || 'recipient'
            }" on the card (or select a contact).`
          );
        }

        const res = await fetch('/api/whatsapp/send', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            to: phone,
            message: action.description || action.title,
            provider: settings.whatsappProvider,
            defaultCountryCode: settings.defaultCountryCode,
            metaAccessToken: settings.metaAccessToken,
            metaPhoneNumberId: settings.metaPhoneNumberId,
            metaTemplateName: settings.metaTemplateName,
            twilioAccountSid: settings.twilioAccountSid,
            twilioAuthToken: settings.twilioAuthToken,
            twilioFromNumber: settings.twilioFromNumber,
          }),
        });

        const data = await res.json();
        if (!res.ok) {
          throw new Error(data.error || 'Failed to send WhatsApp message');
        }

        externalId = data.messageId;
        showToast(
          `WhatsApp sent to ${action.recipient_name || data.to} via ${
            data.provider === 'meta' ? 'Meta Cloud API' : 'Twilio'
          }!`,
          'success'
        );
      } else if (action.action_type === 'outlook_task') {
        const res = await graphService.createOutlookTodoTask(
          action,
          settings.defaultTodoListId
        );
        externalId = res.id;
        externalUrl = res.webUrl;
        showToast(`Created Outlook Task: "${action.title}"`, 'success');
      } else if (action.action_type === 'planner_task') {
        const res = await graphService.createPlannerTask(
          action,
          settings.defaultPlannerPlanId,
          settings.defaultPlannerBucketId
        );
        externalId = res.id;
        externalUrl = res.webUrl;
        showToast(
          `Created Microsoft Planner Task: "${action.title}"`,
          'success'
        );
      } else if (action.action_type === 'outlook_email') {
        const res = await graphService.sendOrDraftOutlookEmail(
          action,
          settings.emailMode
        );
        externalId = res.id;
        externalUrl = res.webUrl;
        showToast(
          res.method === 'graph_send'
            ? `Email sent to ${action.recipient_email}!`
            : res.method === 'graph_draft'
            ? `Email saved to Outlook Drafts!`
            : `Opened Outlook Web Compose for "${action.title}"`,
          'success'
        );
      } else if (action.action_type === 'outlook_event') {
        const res = await graphService.createOutlookCalendarEvent(action);
        externalId = res.id;
        externalUrl = res.webUrl;
        showToast(
          `Scheduled "${action.title}" in Outlook Calendar!`,
          'success'
        );
      }

      const updated = await updateActionRecord(action.id, {
        ...action,
        status: 'completed',
        external_id: externalId,
        external_url: externalUrl,
        error_message: '',
        executed_at: new Date().toISOString(),
      });

      setActions((prev) =>
        prev.map((a) => (a.id === action.id ? updated : a))
      );
    } catch (err: any) {
      const message = err?.message || 'Execution failed';
      showToast(message, 'error');
      try {
        const failed = await updateActionRecord(action.id, {
          ...action,
          status: 'failed',
          error_message: message,
        });
        setActions((prev) =>
          prev.map((a) => (a.id === action.id ? failed : a))
        );
      } catch {
        // ignore
      }
    } finally {
      setExecutingIds((prev) => ({ ...prev, [action.id]: false }));
    }
  };

  // ==========================================================
  // PROCESS VOICE TRANSCRIPT -> EXTRACT & SAVE ACTIONS
  // ==========================================================
  const handleProcessTranscript = async (
    transcript: string,
    durationSeconds: number
  ) => {
    setIsProcessingVoice(true);
    try {
      // 1. Parse transcript into structured actions
      const draftInputs = await parseVoiceTranscriptToActions(
        transcript,
        contacts,
        settings
      );

      // 2. Save Voice Note record in Supabase
      const savedNote = await createVoiceNote({
        transcript,
        summary: `${draftInputs.length} action(s) extracted`,
        duration_seconds: durationSeconds,
        language: settings.speechLanguage,
        ai_provider: settings.aiProvider,
        actions_count: draftInputs.length,
      });

      if (savedNote) {
        setVoiceNotes((prev) => [savedNote, ...prev]);
      }

      // 3. Link voice_note_id and insert actions into Supabase
      const actionsWithNoteId = draftInputs.map((d) => ({
        ...d,
        voice_note_id: savedNote?.id || null,
      }));

      const createdActions = await createActionsBatch(actionsWithNoteId);
      setActions((prev) => [...createdActions, ...prev]);
      setActiveWorkspaceTab('actions');
      setActionFilter('all');

      showToast(
        `Extracted ${createdActions.length} action(s) from your voice note!`,
        'success'
      );

      // 4. Optional Zero-Click Auto-Execute
      if (settings.autoExecuteHighConfidence && createdActions.length > 0) {
        for (const act of createdActions) {
          await handleExecuteAction(act);
        }
      }
    } catch (err: any) {
      showToast(err?.message || 'Could not process voice note', 'error');
    } finally {
      setIsProcessingVoice(false);
    }
  };

  // ==========================================================
  // MANUAL QUICK-ADD ACTION
  // ==========================================================
  const handleQuickAddAction = async (type: ActionType) => {
    const defaultTitles: Record<ActionType, string> = {
      whatsapp: 'New WhatsApp Message',
      outlook_task: 'New Outlook Task',
      planner_task: 'New Planner Task',
      outlook_email: 'Follow-up Email',
      outlook_event: 'Team Sync Meeting',
    };

    try {
      const created = await createActionsBatch([
        {
          action_type: type,
          title: defaultTitles[type],
          description: '',
          priority: 'normal',
          plan_id: settings.defaultPlannerPlanId || undefined,
          plan_name: settings.defaultPlannerPlanName || undefined,
          bucket_id: settings.defaultPlannerBucketId || undefined,
          bucket_name: settings.defaultPlannerBucketName || undefined,
          status: 'draft',
        },
      ]);
      setActions((prev) => [...created, ...prev]);
      setActiveWorkspaceTab('actions');
    } catch (err: any) {
      showToast(err?.message || 'Failed to add action card', 'error');
    }
  };

  const handleUpdateAction = async (
    id: string,
    patch: Partial<AssistantAction>
  ) => {
    try {
      const updated = await updateActionRecord(id, patch);
      setActions((prev) => prev.map((a) => (a.id === id ? updated : a)));
    } catch (err: any) {
      console.warn('Update action error:', err);
    }
  };

  const handleDeleteAction = async (id: string) => {
    try {
      await deleteActionRecord(id);
      setActions((prev) => prev.filter((a) => a.id !== id));
    } catch (err: any) {
      showToast(err?.message || 'Failed to delete action', 'error');
    }
  };

  const handleExecuteAllDrafts = async () => {
    const drafts = actions.filter((a) => a.status === 'draft');
    for (const draft of drafts) {
      await handleExecuteAction(draft);
    }
  };

  // Filtered actions
  const filteredActions = actions.filter((a) => {
    if (actionFilter === 'all') return true;
    if (actionFilter === 'draft') return a.status === 'draft' || a.status === 'failed';
    if (actionFilter === 'completed') return a.status === 'completed';
    return a.action_type === actionFilter;
  });

  const pendingCount = actions.filter(
    (a) => a.status === 'draft' || a.status === 'failed'
  ).length;
  const whatsappSentCount = actions.filter(
    (a) => a.action_type === 'whatsapp' && a.status === 'completed'
  ).length;
  const outlookCount = actions.filter(
    (a) =>
      (a.action_type === 'outlook_task' ||
        a.action_type === 'outlook_email' ||
        a.action_type === 'outlook_event') &&
      a.status === 'completed'
  ).length;
  const plannerCount = actions.filter(
    (a) => a.action_type === 'planner_task' && a.status === 'completed'
  ).length;

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100">
      {/* Top Toast Notification */}
      {toast && (
        <div className="fixed bottom-5 right-5 z-50 max-w-md rounded-xl border px-4 py-3 text-xs font-medium shadow-2xl backdrop-blur-md transition-all bg-slate-900/95 border-slate-700 text-white flex items-center gap-2.5">
          <span
            className={`h-2.5 w-2.5 rounded-full shrink-0 ${
              toast.type === 'success'
                ? 'bg-emerald-400'
                : toast.type === 'error'
                ? 'bg-rose-400'
                : 'bg-indigo-400'
            }`}
          />
          <span>{toast.text}</span>
        </div>
      )}

      {/* Top Navigation Bar */}
      <header className="sticky top-0 z-30 border-b border-slate-800/80 bg-slate-950/90 backdrop-blur-md">
        <div className="mx-auto flex max-w-7xl flex-wrap items-center justify-between gap-4 px-4 py-3.5 sm:px-6">
          <div className="flex items-center gap-3">
            <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-gradient-to-br from-indigo-500 to-purple-600 text-white shadow-lg shadow-indigo-500/20">
              <Sparkles className="h-5 w-5" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-lg font-bold tracking-tight text-white">
                  Protemp Personal Assistant
                </h1>
                <span className="rounded-full bg-emerald-500/15 border border-emerald-500/30 px-2 py-0.5 text-[10px] font-semibold text-emerald-300">
                  Supabase Live
                </span>
              </div>
              <p className="text-xs text-slate-400">
                Voice-to-WhatsApp Cloud API &bull; Microsoft Outlook Tasks &amp; Email &bull; Microsoft Planner
              </p>
            </div>
          </div>

          {/* Right Header Actions */}
          <div className="flex flex-wrap items-center gap-2.5">
            {/* Contacts Directory Button */}
            <button
              type="button"
              onClick={() => setIsContactsOpen(true)}
              className="inline-flex items-center gap-1.5 rounded-xl border border-slate-800 bg-slate-900 hover:bg-slate-800 px-3.5 py-2 text-xs font-medium text-slate-200 transition"
            >
              <Users className="h-4 w-4 text-emerald-400" />
              Contacts ({contacts.length})
            </button>

            {/* Microsoft 365 Connection Button */}
            {m365User ? (
              <div className="flex items-center gap-1.5 rounded-xl border border-sky-500/30 bg-sky-500/10 px-3 py-1.5 text-xs text-sky-200">
                <span className="h-2 w-2 rounded-full bg-emerald-400" />
                <span className="max-w-[160px] truncate font-medium">
                  {m365User.name}
                </span>
                <button
                  type="button"
                  onClick={async () => {
                    await graphService.logout();
                    setM365User(null);
                    showToast('Signed out of Microsoft 365', 'info');
                  }}
                  className="ml-1 rounded p-1 text-slate-400 hover:bg-slate-800 hover:text-rose-300"
                  title="Sign out of Microsoft 365"
                >
                  <LogOut className="h-3.5 w-3.5" />
                </button>
              </div>
            ) : (
              <button
                type="button"
                onClick={async () => {
                  try {
                    await graphService.login();
                  } catch (err: any) {
                    showToast(err?.message || 'Microsoft 365 sign-in failed', 'error');
                  }
                }}
                className="inline-flex items-center gap-1.5 rounded-xl bg-sky-600 hover:bg-sky-500 px-3.5 py-2 text-xs font-semibold text-white shadow-sm transition"
              >
                <LogIn className="h-4 w-4" />
                Connect Microsoft 365
              </button>
            )}

            {/* Settings Modal Button */}
            <button
              type="button"
              onClick={() => setIsSettingsOpen(true)}
              className="inline-flex items-center gap-1.5 rounded-xl border border-slate-800 bg-slate-900 hover:bg-slate-800 px-3.5 py-2 text-xs font-medium text-slate-200 transition"
            >
              <Settings className="h-4 w-4 text-indigo-400" />
              API Settings
            </button>
          </div>
        </div>
      </header>

      {/* Main Content Container */}
      <main className="mx-auto max-w-7xl px-4 py-6 sm:px-6 space-y-6">
        {/* KPI Summary Row */}
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
          <div className="rounded-2xl border border-slate-800/90 bg-slate-900/70 p-4 flex items-center justify-between">
            <div>
              <div className="text-xs text-slate-400">Pending Draft Actions</div>
              <div className="mt-1 text-2xl font-bold text-white">
                {pendingCount}
              </div>
            </div>
            <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-amber-500/15 text-amber-400">
              <Clock className="h-5 w-5" />
            </div>
          </div>

          <div className="rounded-2xl border border-slate-800/90 bg-slate-900/70 p-4 flex items-center justify-between">
            <div>
              <div className="text-xs text-slate-400">WhatsApps Sent (API)</div>
              <div className="mt-1 text-2xl font-bold text-emerald-400">
                {whatsappSentCount}
              </div>
            </div>
            <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-emerald-500/15 text-emerald-400">
              <MessageSquare className="h-5 w-5" />
            </div>
          </div>

          <div className="rounded-2xl border border-slate-800/90 bg-slate-900/70 p-4 flex items-center justify-between">
            <div>
              <div className="text-xs text-slate-400">
                Outlook Tasks &amp; Emails
              </div>
              <div className="mt-1 text-2xl font-bold text-sky-400">
                {outlookCount}
              </div>
            </div>
            <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-sky-500/15 text-sky-400">
              <CheckSquare className="h-5 w-5" />
            </div>
          </div>

          <div className="rounded-2xl border border-slate-800/90 bg-slate-900/70 p-4 flex items-center justify-between">
            <div>
              <div className="text-xs text-slate-400">
                Planner Tasks Generated
              </div>
              <div className="mt-1 text-2xl font-bold text-purple-400">
                {plannerCount}
              </div>
            </div>
            <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-purple-500/15 text-purple-400">
              <LayoutGrid className="h-5 w-5" />
            </div>
          </div>
        </div>

        {/* Voice Recorder & Multi-Action Extractor */}
        <VoiceRecorderCard
          language={settings.speechLanguage}
          aiProvider={settings.aiProvider}
          isProcessing={isProcessingVoice}
          onProcessTranscript={handleProcessTranscript}
        />

        {/* Workspace Navigation Tabs & Quick Add Bar */}
        <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-800 pb-3">
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => setActiveWorkspaceTab('actions')}
              className={`rounded-xl px-4 py-2 text-xs font-semibold transition ${
                activeWorkspaceTab === 'actions'
                  ? 'bg-indigo-600 text-white shadow-sm'
                  : 'bg-slate-900 text-slate-400 hover:text-white'
              }`}
            >
              Action Queue &amp; History ({actions.length})
            </button>
            <button
              type="button"
              onClick={() => setActiveWorkspaceTab('live_m365')}
              className={`rounded-xl px-4 py-2 text-xs font-semibold transition ${
                activeWorkspaceTab === 'live_m365'
                  ? 'bg-sky-600 text-white shadow-sm'
                  : 'bg-slate-900 text-slate-400 hover:text-white'
              }`}
            >
              Live Outlook &amp; Planner Tasks
            </button>
            <button
              type="button"
              onClick={() => setActiveWorkspaceTab('transcripts')}
              className={`rounded-xl px-4 py-2 text-xs font-semibold transition ${
                activeWorkspaceTab === 'transcripts'
                  ? 'bg-purple-600 text-white shadow-sm'
                  : 'bg-slate-900 text-slate-400 hover:text-white'
              }`}
            >
              Voice Note Log ({voiceNotes.length})
            </button>
          </div>

          {/* Quick Manual Add Buttons */}
          <div className="flex flex-wrap items-center gap-2">
            <span className="text-[11px] text-slate-500 mr-1">Manual Add:</span>
            <button
              type="button"
              onClick={() => handleQuickAddAction('whatsapp')}
              className="inline-flex items-center gap-1 rounded-lg border border-emerald-500/30 bg-emerald-500/10 hover:bg-emerald-500/20 px-2.5 py-1.5 text-xs font-medium text-emerald-300 transition"
            >
              <Plus className="h-3.5 w-3.5" /> WhatsApp
            </button>
            <button
              type="button"
              onClick={() => handleQuickAddAction('outlook_task')}
              className="inline-flex items-center gap-1 rounded-lg border border-sky-500/30 bg-sky-500/10 hover:bg-sky-500/20 px-2.5 py-1.5 text-xs font-medium text-sky-300 transition"
            >
              <Plus className="h-3.5 w-3.5" /> Outlook Task
            </button>
            <button
              type="button"
              onClick={() => handleQuickAddAction('planner_task')}
              className="inline-flex items-center gap-1 rounded-lg border border-purple-500/30 bg-purple-500/10 hover:bg-purple-500/20 px-2.5 py-1.5 text-xs font-medium text-purple-300 transition"
            >
              <Plus className="h-3.5 w-3.5" /> Planner Task
            </button>
            <button
              type="button"
              onClick={() => handleQuickAddAction('outlook_email')}
              className="inline-flex items-center gap-1 rounded-lg border border-amber-500/30 bg-amber-500/10 hover:bg-amber-500/20 px-2.5 py-1.5 text-xs font-medium text-amber-300 transition"
            >
              <Plus className="h-3.5 w-3.5" /> Email
            </button>
          </div>
        </div>

        {/* TAB 1: ACTION QUEUE & HISTORY */}
        {activeWorkspaceTab === 'actions' && (
          <div className="space-y-4">
            {/* Filter Pills & Batch Execute */}
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div className="flex flex-wrap items-center gap-1.5">
                {[
                  { id: 'all', label: 'All Actions' },
                  { id: 'draft', label: `Pending (${pendingCount})` },
                  { id: 'whatsapp', label: 'WhatsApp' },
                  { id: 'outlook_task', label: 'Outlook Tasks' },
                  { id: 'planner_task', label: 'MS Planner' },
                  { id: 'outlook_email', label: 'Outlook Email' },
                  { id: 'completed', label: 'Completed' },
                ].map((tab) => (
                  <button
                    key={tab.id}
                    type="button"
                    onClick={() => setActionFilter(tab.id as any)}
                    className={`rounded-lg px-3 py-1.5 text-xs font-medium transition ${
                      actionFilter === tab.id
                        ? 'bg-slate-800 text-white border border-slate-600'
                        : 'bg-slate-900/60 text-slate-400 hover:bg-slate-900 hover:text-slate-200'
                    }`}
                  >
                    {tab.label}
                  </button>
                ))}
              </div>

              {pendingCount > 0 && (
                <button
                  type="button"
                  onClick={handleExecuteAllDrafts}
                  className="inline-flex items-center gap-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 px-4 py-2 text-xs font-semibold text-white shadow-md transition"
                >
                  <Play className="h-3.5 w-3.5 fill-current" />
                  Execute All Pending ({pendingCount})
                </button>
              )}
            </div>

            {filteredActions.length === 0 ? (
              <div className="rounded-2xl border border-dashed border-slate-800 bg-slate-900/30 py-14 text-center">
                <Sparkles className="mx-auto h-8 w-8 text-slate-600 mb-2" />
                <h3 className="text-sm font-semibold text-slate-300">
                  No actions in this view yet
                </h3>
                <p className="mt-1 text-xs text-slate-500 max-w-md mx-auto">
                  Record a voice note above (or click one of the example voice commands) to automatically generate WhatsApp messages, Outlook tasks, and Microsoft Planner tasks.
                </p>
              </div>
            ) : (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {filteredActions.map((act) => (
                  <ActionCard
                    key={act.id}
                    action={act}
                    contacts={contacts}
                    plannerPlans={plannerPlans}
                    bucketsByPlan={bucketsByPlan}
                    onLoadBuckets={handleLoadBuckets}
                    onUpdate={handleUpdateAction}
                    onExecute={handleExecuteAction}
                    onDelete={handleDeleteAction}
                    isExecuting={Boolean(executingIds[act.id])}
                  />
                ))}
              </div>
            )}
          </div>
        )}

        {/* TAB 2: LIVE MICROSOFT 365 TASKS (OUTLOOK TO DO & PLANNER) */}
        {activeWorkspaceTab === 'live_m365' && (
          <div className="space-y-4">
            {!m365User ? (
              <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-10 text-center">
                <LayoutGrid className="mx-auto h-10 w-10 text-sky-400 mb-3" />
                <h3 className="text-base font-semibold text-white">
                  Connect Microsoft 365 to View Live Outlook &amp; Planner Tasks
                </h3>
                <p className="mt-1 text-xs text-slate-400 max-w-md mx-auto mb-5">
                  Sign in with your Protemp Microsoft 365 account to view and complete your live Outlook To Do lists and Microsoft Planner tasks directly from your assistant.
                </p>
                <button
                  type="button"
                  onClick={() => graphService.login()}
                  className="inline-flex items-center gap-2 rounded-xl bg-sky-600 hover:bg-sky-500 px-5 py-2.5 text-xs font-semibold text-white shadow-lg"
                >
                  <LogIn className="h-4 w-4" /> Connect Microsoft 365
                </button>
              </div>
            ) : (
              <div className="space-y-4">
                <div className="flex items-center justify-between">
                  <p className="text-xs text-slate-400">
                    Showing live active tasks synced from your Microsoft 365 account ({m365User.email})
                  </p>
                  <button
                    type="button"
                    onClick={refreshLiveM365Tasks}
                    disabled={isLoadingLiveM365}
                    className="inline-flex items-center gap-1.5 rounded-lg border border-slate-800 bg-slate-900 hover:bg-slate-800 px-3 py-1.5 text-xs text-slate-200"
                  >
                    <RefreshCw
                      className={`h-3.5 w-3.5 ${
                        isLoadingLiveM365 ? 'animate-spin' : ''
                      }`}
                    />
                    Refresh M365
                  </button>
                </div>

                <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                  {/* Column 1: Live Outlook To Do Tasks */}
                  <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-4 space-y-3">
                    <div className="flex items-center justify-between border-b border-slate-800 pb-3">
                      <div className="flex items-center gap-2">
                        <CheckSquare className="h-4 w-4 text-sky-400" />
                        <h3 className="text-sm font-semibold text-white">
                          Microsoft To Do / Outlook Tasks ({liveTodoTasks.length})
                        </h3>
                      </div>
                      <a
                        href="https://to-do.office.com/tasks/"
                        target="_blank"
                        rel="noopener noreferrer"
                        className="text-xs text-sky-400 hover:underline inline-flex items-center gap-1"
                      >
                        Open To Do <ExternalLink className="h-3 w-3" />
                      </a>
                    </div>

                    {liveTodoTasks.length === 0 ? (
                      <div className="py-8 text-center text-xs text-slate-500">
                        {isLoadingLiveM365
                          ? 'Loading Outlook tasks...'
                          : 'No open tasks found in Outlook To Do.'}
                      </div>
                    ) : (
                      <div className="divide-y divide-slate-800/80 max-h-96 overflow-y-auto">
                        {liveTodoTasks.map((t) => (
                          <div
                            key={t.id}
                            className="flex items-center justify-between gap-3 py-2.5 px-2 hover:bg-slate-800/40 rounded-lg"
                          >
                            <div>
                              <div className="text-xs font-medium text-white">
                                {t.title}
                              </div>
                              <div className="text-[11px] text-slate-400">
                                List: {t.listName}
                                {t.dueDateTime?.dateTime
                                  ? ` • Due ${t.dueDateTime.dateTime.split('T')[0]}`
                                  : ''}
                              </div>
                            </div>
                            <button
                              type="button"
                              onClick={async () => {
                                await graphService.completeOutlookTask(
                                  t.listId,
                                  t.id
                                );
                                setLiveTodoTasks((prev) =>
                                  prev.filter((item) => item.id !== t.id)
                                );
                                showToast(
                                  `Completed Outlook task: "${t.title}"`,
                                  'success'
                                );
                              }}
                              className="inline-flex items-center gap-1 rounded-lg border border-emerald-500/30 bg-emerald-500/10 hover:bg-emerald-500/20 px-2.5 py-1 text-[11px] font-medium text-emerald-300"
                            >
                              <CheckCircle2 className="h-3.5 w-3.5" /> Done
                            </button>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>

                  {/* Column 2: Live Microsoft Planner Tasks */}
                  <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-4 space-y-3">
                    <div className="flex items-center justify-between border-b border-slate-800 pb-3">
                      <div className="flex items-center gap-2">
                        <LayoutGrid className="h-4 w-4 text-purple-400" />
                        <h3 className="text-sm font-semibold text-white">
                          Microsoft Planner Tasks ({livePlannerTasks.length})
                        </h3>
                      </div>
                      <a
                        href="https://tasks.office.com"
                        target="_blank"
                        rel="noopener noreferrer"
                        className="text-xs text-purple-400 hover:underline inline-flex items-center gap-1"
                      >
                        Open Planner <ExternalLink className="h-3 w-3" />
                      </a>
                    </div>

                    {livePlannerTasks.length === 0 ? (
                      <div className="py-8 text-center text-xs text-slate-500">
                        {isLoadingLiveM365
                          ? 'Loading Planner tasks...'
                          : 'No active Planner tasks assigned to you.'}
                      </div>
                    ) : (
                      <div className="divide-y divide-slate-800/80 max-h-96 overflow-y-auto">
                        {livePlannerTasks.map((pt) => (
                          <div
                            key={pt.id}
                            className="flex items-center justify-between gap-3 py-2.5 px-2 hover:bg-slate-800/40 rounded-lg"
                          >
                            <div>
                              <div className="text-xs font-medium text-white">
                                {pt.title}
                              </div>
                              <div className="text-[11px] text-slate-400">
                                {pt.dueDateTime
                                  ? `Due ${pt.dueDateTime.split('T')[0]}`
                                  : 'No due date'}
                              </div>
                            </div>
                            <button
                              type="button"
                              onClick={async () => {
                                await graphService.completePlannerTask(pt.id);
                                setLivePlannerTasks((prev) =>
                                  prev.filter((item) => item.id !== pt.id)
                                );
                                showToast(
                                  `Completed Planner task: "${pt.title}"`,
                                  'success'
                                );
                              }}
                              className="inline-flex items-center gap-1 rounded-lg border border-emerald-500/30 bg-emerald-500/10 hover:bg-emerald-500/20 px-2.5 py-1 text-[11px] font-medium text-emerald-300"
                            >
                              <CheckCircle2 className="h-3.5 w-3.5" /> Complete
                            </button>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        {/* TAB 3: VOICE NOTE LOG */}
        {activeWorkspaceTab === 'transcripts' && (
          <div className="space-y-3">
            {voiceNotes.length === 0 ? (
              <div className="rounded-2xl border border-dashed border-slate-800 py-12 text-center text-xs text-slate-500">
                No voice notes recorded yet.
              </div>
            ) : (
              voiceNotes.map((vn) => (
                <div
                  key={vn.id}
                  className="flex flex-wrap items-center justify-between gap-4 rounded-xl border border-slate-800 bg-slate-900/70 p-4"
                >
                  <div className="flex items-start gap-3">
                    <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-purple-500/15 text-purple-400">
                      <FileAudio className="h-4 w-4" />
                    </div>
                    <div>
                      <p className="text-sm text-white font-medium">
                        &ldquo;{vn.transcript}&rdquo;
                      </p>
                      <div className="mt-1 flex items-center gap-3 text-[11px] text-slate-400">
                        <span>
                          {new Date(vn.created_at).toLocaleString()}
                        </span>
                        <span>&bull;</span>
                        <span>{vn.actions_count} action(s) extracted</span>
                        <span>&bull;</span>
                        <span>{vn.language}</span>
                      </div>
                    </div>
                  </div>

                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() =>
                        handleProcessTranscript(
                          vn.transcript,
                          vn.duration_seconds
                        )
                      }
                      className="rounded-lg border border-slate-700 bg-slate-800 hover:bg-slate-700 px-3 py-1.5 text-xs text-slate-200 transition"
                    >
                      Re-Extract Actions
                    </button>
                    <button
                      type="button"
                      onClick={async () => {
                        await deleteVoiceNoteById(vn.id);
                        setVoiceNotes((prev) =>
                          prev.filter((item) => item.id !== vn.id)
                        );
                      }}
                      className="rounded-lg p-1.5 text-slate-500 hover:bg-slate-800 hover:text-rose-400"
                    >
                      <Trash2 className="h-4 w-4" />
                    </button>
                  </div>
                </div>
              ))
            )}
          </div>
        )}
      </main>

      {/* Contacts Directory Modal */}
      <ContactsModal
        isOpen={isContactsOpen}
        onClose={() => setIsContactsOpen(false)}
        contacts={contacts}
        onAddOrUpdate={async (contactInput) => {
          const saved = await upsertContact(contactInput);
          if (saved) {
            setContacts((prev) => {
              const exists = prev.some((c) => c.id === saved.id);
              return exists
                ? prev.map((c) => (c.id === saved.id ? saved : c))
                : [...prev, saved].sort((a, b) => a.name.localeCompare(b.name));
            });
            showToast(`Saved contact "${saved.name}"`, 'success');
          }
        }}
        onDelete={async (id) => {
          await deleteContactById(id);
          setContacts((prev) => prev.filter((c) => c.id !== id));
        }}
        onSyncM365Contacts={async () => {
          const people = await graphService.fetchMicrosoftPeople();
          let added = 0;
          for (const p of people) {
            if (!p.name) continue;
            const already = contacts.some(
              (c) => c.name.toLowerCase() === p.name!.toLowerCase()
            );
            if (!already) {
              const saved = await upsertContact({
                name: p.name,
                email: p.email,
                phone: p.phone,
                role: p.role,
                company: p.company || 'Protemp',
              });
              if (saved) {
                added++;
                setContacts((prev) => [...prev, saved]);
              }
            }
          }
          return added;
        }}
        isM365SignedIn={Boolean(m365User)}
      />

      {/* Settings Modal */}
      <SettingsModal
        isOpen={isSettingsOpen}
        onClose={() => setIsSettingsOpen(false)}
        settings={settings}
        onSave={async (nextSettings) => {
          setSettings(nextSettings);
          setStoredAzureConfig(
            nextSettings.azureClientId,
            nextSettings.azureTenantId
          );
          await saveAssistantSettings(nextSettings);
          showToast('Assistant settings saved!', 'success');
        }}
        plannerPlans={plannerPlans}
        todoLists={todoLists}
        bucketsByPlan={bucketsByPlan}
        onLoadBuckets={handleLoadBuckets}
      />
    </div>
  );
}
