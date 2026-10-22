'use client';

import React, { useState, useEffect } from 'react';
import {
  AssistantAction,
  Contact,
  PlannerBucket,
  PlannerPlan,
  ActionType,
  ActionPriority,
} from '@/lib/types';
import {
  MessageSquare,
  CheckSquare,
  LayoutGrid,
  Mail,
  Calendar,
  Send,
  Trash2,
  CheckCircle2,
  AlertCircle,
  ExternalLink,
  Loader2,
  User,
  Phone,
  Clock,
  Flag,
} from 'lucide-react';

interface ActionCardProps {
  action: AssistantAction;
  contacts: Contact[];
  plannerPlans: PlannerPlan[];
  bucketsByPlan: Record<string, PlannerBucket[]>;
  onLoadBuckets: (planId: string) => Promise<void>;
  onUpdate: (id: string, patch: Partial<AssistantAction>) => Promise<void>;
  onExecute: (action: AssistantAction) => Promise<void>;
  onDelete: (id: string) => Promise<void>;
  isExecuting: boolean;
}

const ACTION_META: Record<
  ActionType,
  {
    label: string;
    badgeClass: string;
    borderClass: string;
    buttonLabel: string;
    buttonClass: string;
    icon: React.ComponentType<{ className?: string }>;
  }
> = {
  whatsapp: {
    label: 'WhatsApp API Message',
    badgeClass: 'bg-emerald-500/15 text-emerald-300 border-emerald-500/30',
    borderClass: 'border-emerald-500/30 hover:border-emerald-500/50',
    buttonLabel: 'Send WhatsApp Now',
    buttonClass: 'bg-emerald-600 hover:bg-emerald-500 text-white',
    icon: MessageSquare,
  },
  outlook_task: {
    label: 'Outlook / Microsoft To Do Task',
    badgeClass: 'bg-sky-500/15 text-sky-300 border-sky-500/30',
    borderClass: 'border-sky-500/30 hover:border-sky-500/50',
    buttonLabel: 'Create in Outlook Tasks',
    buttonClass: 'bg-sky-600 hover:bg-sky-500 text-white',
    icon: CheckSquare,
  },
  planner_task: {
    label: 'Microsoft Planner Task',
    badgeClass: 'bg-purple-500/15 text-purple-300 border-purple-500/30',
    borderClass: 'border-purple-500/30 hover:border-purple-500/50',
    buttonLabel: 'Create in MS Planner',
    buttonClass: 'bg-purple-600 hover:bg-purple-500 text-white',
    icon: LayoutGrid,
  },
  outlook_email: {
    label: 'Outlook Email',
    badgeClass: 'bg-amber-500/15 text-amber-300 border-amber-500/30',
    borderClass: 'border-amber-500/30 hover:border-amber-500/50',
    buttonLabel: 'Send via Outlook',
    buttonClass: 'bg-amber-600 hover:bg-amber-500 text-white',
    icon: Mail,
  },
  outlook_event: {
    label: 'Outlook Calendar Event',
    badgeClass: 'bg-indigo-500/15 text-indigo-300 border-indigo-500/30',
    borderClass: 'border-indigo-500/30 hover:border-indigo-500/50',
    buttonLabel: 'Schedule in Outlook',
    buttonClass: 'bg-indigo-600 hover:bg-indigo-500 text-white',
    icon: Calendar,
  },
};

export default function ActionCard({
  action,
  contacts,
  plannerPlans,
  bucketsByPlan,
  onLoadBuckets,
  onUpdate,
  onExecute,
  onDelete,
  isExecuting,
}: ActionCardProps) {
  const [title, setTitle] = useState(action.title);
  const [description, setDescription] = useState(action.description || '');
  const [recipientName, setRecipientName] = useState(
    action.recipient_name || ''
  );
  const [recipientPhone, setRecipientPhone] = useState(
    action.recipient_phone || ''
  );
  const [recipientEmail, setRecipientEmail] = useState(
    action.recipient_email || ''
  );
  const [dueDate, setDueDate] = useState(
    action.due_date ? action.due_date.split('T')[0] : ''
  );
  const [startTime, setStartTime] = useState(action.start_time || '09:00');
  const [priority, setPriority] = useState<ActionPriority>(
    action.priority || 'normal'
  );
  const [actionType, setActionType] = useState<ActionType>(action.action_type);
  const [planId, setPlanId] = useState(action.plan_id || '');
  const [bucketId, setBucketId] = useState(action.bucket_id || '');

  useEffect(() => {
    if (actionType === 'planner_task' && planId) {
      onLoadBuckets(planId);
    }
  }, [actionType, planId]);

  const meta = ACTION_META[actionType] || ACTION_META.outlook_task;
  const Icon = meta.icon;
  const isCompleted = action.status === 'completed';

  const handleContactSelect = (contactId: string) => {
    const found = contacts.find((c) => c.id === contactId);
    if (!found) return;
    setRecipientName(found.name);
    if (found.phone) setRecipientPhone(found.phone);
    if (found.email) setRecipientEmail(found.email);
    onUpdate(action.id, {
      recipient_name: found.name,
      recipient_phone: found.phone || recipientPhone,
      recipient_email: found.email || recipientEmail,
    });
  };

  const buildCurrentAction = (): AssistantAction => ({
    ...action,
    action_type: actionType,
    title,
    description,
    recipient_name: recipientName,
    recipient_phone: recipientPhone,
    recipient_email: recipientEmail,
    due_date: dueDate || undefined,
    start_time: startTime || undefined,
    priority,
    plan_id: planId || undefined,
    bucket_id: bucketId || undefined,
  });

  const handleSaveField = (patch: Partial<AssistantAction>) => {
    onUpdate(action.id, patch);
  };

  const currentBuckets = planId ? bucketsByPlan[planId] || [] : [];

  return (
    <div
      className={`rounded-xl border bg-slate-900/90 p-4 shadow-md transition-all ${
        isCompleted
          ? 'border-emerald-500/40 bg-emerald-950/10 opacity-85'
          : meta.borderClass
      }`}
    >
      {/* Card Header */}
      <div className="flex flex-wrap items-center justify-between gap-2 mb-3">
        <div className="flex items-center gap-2">
          <span
            className={`inline-flex items-center gap-1.5 rounded-lg border px-2.5 py-1 text-xs font-medium ${meta.badgeClass}`}
          >
            <Icon className="h-3.5 w-3.5" />
            {meta.label}
          </span>

          {/* Switch Action Type Dropdown if not completed */}
          {!isCompleted && (
            <select
              value={actionType}
              onChange={(e) => {
                const nextType = e.target.value as ActionType;
                setActionType(nextType);
                handleSaveField({ action_type: nextType });
              }}
              className="rounded-lg border border-slate-700 bg-slate-800 px-2 py-1 text-[11px] text-slate-300 focus:border-indigo-500 focus:outline-none"
              title="Change Action Type"
            >
              <option value="whatsapp">WhatsApp API</option>
              <option value="outlook_task">Outlook Task (To Do)</option>
              <option value="planner_task">Microsoft Planner</option>
              <option value="outlook_email">Outlook Email</option>
              <option value="outlook_event">Outlook Calendar Event</option>
            </select>
          )}
        </div>

        <div className="flex items-center gap-2">
          {isCompleted && (
            <span className="inline-flex items-center gap-1 rounded-full bg-emerald-500/20 px-2.5 py-0.5 text-xs font-medium text-emerald-300">
              <CheckCircle2 className="h-3.5 w-3.5" /> Sent / Created
            </span>
          )}
          {action.status === 'failed' && (
            <span className="inline-flex items-center gap-1 rounded-full bg-rose-500/20 px-2.5 py-0.5 text-xs font-medium text-rose-300">
              <AlertCircle className="h-3.5 w-3.5" /> Action Needed
            </span>
          )}
          <button
            type="button"
            onClick={() => onDelete(action.id)}
            className="rounded-lg p-1.5 text-slate-500 hover:bg-slate-800 hover:text-rose-400 transition"
            title="Delete Action"
          >
            <Trash2 className="h-4 w-4" />
          </button>
        </div>
      </div>

      {/* Main Editable Fields */}
      <div className="space-y-3">
        {/* Title / Subject */}
        <div>
          <label className="block text-[11px] font-medium text-slate-400 mb-1">
            {actionType === 'whatsapp'
              ? 'Action Summary'
              : actionType === 'outlook_email'
              ? 'Email Subject'
              : 'Task / Event Title'}
          </label>
          <input
            type="text"
            value={title}
            disabled={isCompleted}
            onChange={(e) => setTitle(e.target.value)}
            onBlur={() => handleSaveField({ title })}
            className="w-full rounded-lg border border-slate-800 bg-slate-950/80 px-3 py-1.5 text-sm font-medium text-white focus:border-indigo-500 focus:outline-none disabled:opacity-60"
          />
        </div>

        {/* WhatsApp Recipient & Phone Row */}
        {actionType === 'whatsapp' && (
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
            <div>
              <label className="block text-[11px] font-medium text-slate-400 mb-1">
                Quick Pick Contact
              </label>
              <select
                disabled={isCompleted}
                onChange={(e) => handleContactSelect(e.target.value)}
                defaultValue=""
                className="w-full rounded-lg border border-slate-800 bg-slate-950/80 px-2.5 py-1.5 text-xs text-slate-200 focus:border-emerald-500 focus:outline-none"
              >
                <option value="">-- Choose from Contacts --</option>
                {contacts.map((c) => (
                  <option key={c.id} value={c.id}>
                    {c.name} {c.phone ? `(${c.phone})` : ''}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-[11px] font-medium text-slate-400 mb-1 flex items-center gap-1">
                <User className="h-3 w-3" /> Recipient Name
              </label>
              <input
                type="text"
                value={recipientName}
                disabled={isCompleted}
                onChange={(e) => setRecipientName(e.target.value)}
                onBlur={() => handleSaveField({ recipient_name: recipientName })}
                placeholder="e.g. John"
                className="w-full rounded-lg border border-slate-800 bg-slate-950/80 px-2.5 py-1.5 text-xs text-white focus:border-emerald-500 focus:outline-none"
              />
            </div>

            <div>
              <label className="block text-[11px] font-medium text-slate-400 mb-1 flex items-center gap-1">
                <Phone className="h-3 w-3" /> WhatsApp Number (E.164)
              </label>
              <input
                type="text"
                value={recipientPhone}
                disabled={isCompleted}
                onChange={(e) => setRecipientPhone(e.target.value)}
                onBlur={() =>
                  handleSaveField({ recipient_phone: recipientPhone })
                }
                placeholder="e.g. +27821234567"
                className="w-full rounded-lg border border-slate-800 bg-slate-950/80 px-2.5 py-1.5 text-xs text-white focus:border-emerald-500 focus:outline-none"
              />
            </div>
          </div>
        )}

        {/* Email / Calendar Recipient Row */}
        {(actionType === 'outlook_email' || actionType === 'outlook_event') && (
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
            <div>
              <label className="block text-[11px] font-medium text-slate-400 mb-1">
                Quick Pick Contact
              </label>
              <select
                disabled={isCompleted}
                onChange={(e) => handleContactSelect(e.target.value)}
                defaultValue=""
                className="w-full rounded-lg border border-slate-800 bg-slate-950/80 px-2.5 py-1.5 text-xs text-slate-200 focus:border-amber-500 focus:outline-none"
              >
                <option value="">-- Choose from Contacts --</option>
                {contacts.map((c) => (
                  <option key={c.id} value={c.id}>
                    {c.name} {c.email ? `(${c.email})` : ''}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-[11px] font-medium text-slate-400 mb-1">
                Recipient Name
              </label>
              <input
                type="text"
                value={recipientName}
                disabled={isCompleted}
                onChange={(e) => setRecipientName(e.target.value)}
                onBlur={() => handleSaveField({ recipient_name: recipientName })}
                placeholder="Name"
                className="w-full rounded-lg border border-slate-800 bg-slate-950/80 px-2.5 py-1.5 text-xs text-white focus:border-amber-500 focus:outline-none"
              />
            </div>

            <div>
              <label className="block text-[11px] font-medium text-slate-400 mb-1">
                Email Address
              </label>
              <input
                type="email"
                value={recipientEmail}
                disabled={isCompleted}
                onChange={(e) => setRecipientEmail(e.target.value)}
                onBlur={() =>
                  handleSaveField({ recipient_email: recipientEmail })
                }
                placeholder="colleague@protemp.co.za"
                className="w-full rounded-lg border border-slate-800 bg-slate-950/80 px-2.5 py-1.5 text-xs text-white focus:border-amber-500 focus:outline-none"
              />
            </div>
          </div>
        )}

        {/* Microsoft Planner Plan & Bucket Selector */}
        {actionType === 'planner_task' && (
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
            <div>
              <label className="block text-[11px] font-medium text-slate-400 mb-1">
                Planner Plan
              </label>
              <select
                value={planId}
                disabled={isCompleted}
                onChange={(e) => {
                  const nextPlanId = e.target.value;
                  const planObj = plannerPlans.find(
                    (p) => p.id === nextPlanId
                  );
                  setPlanId(nextPlanId);
                  setBucketId('');
                  handleSaveField({
                    plan_id: nextPlanId,
                    plan_name: planObj?.title || '',
                    bucket_id: '',
                  });
                }}
                className="w-full rounded-lg border border-slate-800 bg-slate-950/80 px-2.5 py-1.5 text-xs text-slate-200 focus:border-purple-500 focus:outline-none"
              >
                <option value="">
                  {plannerPlans.length > 0
                    ? '-- Auto / Select Planner Plan --'
                    : 'Connect Microsoft 365 to load Plans'}
                </option>
                {plannerPlans.map((p) => (
                  <option key={p.id} value={p.id}>
                    {p.title}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-[11px] font-medium text-slate-400 mb-1">
                Planner Bucket
              </label>
              <select
                value={bucketId}
                disabled={isCompleted || !planId}
                onChange={(e) => {
                  const nextBucketId = e.target.value;
                  const bObj = currentBuckets.find(
                    (b) => b.id === nextBucketId
                  );
                  setBucketId(nextBucketId);
                  handleSaveField({
                    bucket_id: nextBucketId,
                    bucket_name: bObj?.name || '',
                  });
                }}
                className="w-full rounded-lg border border-slate-800 bg-slate-950/80 px-2.5 py-1.5 text-xs text-slate-200 focus:border-purple-500 focus:outline-none disabled:opacity-50"
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
        )}

        {/* Message Body / Task Description */}
        <div>
          <label className="block text-[11px] font-medium text-slate-400 mb-1">
            {actionType === 'whatsapp'
              ? 'WhatsApp Message Text'
              : actionType === 'outlook_email'
              ? 'Email Message Body'
              : 'Notes / Details'}
          </label>
          <textarea
            rows={2}
            value={description}
            disabled={isCompleted}
            onChange={(e) => setDescription(e.target.value)}
            onBlur={() => handleSaveField({ description })}
            className="w-full rounded-lg border border-slate-800 bg-slate-950/80 p-2.5 text-xs text-slate-200 focus:border-indigo-500 focus:outline-none resize-none disabled:opacity-60"
          />
        </div>

        {/* Date, Time & Priority Row */}
        {actionType !== 'whatsapp' && (
          <div className="grid grid-cols-2 sm:grid-cols-3 gap-2.5">
            <div>
              <label className="block text-[11px] font-medium text-slate-400 mb-1 flex items-center gap-1">
                <Calendar className="h-3 w-3" /> Due / Event Date
              </label>
              <input
                type="date"
                value={dueDate}
                disabled={isCompleted}
                onChange={(e) => {
                  setDueDate(e.target.value);
                  handleSaveField({ due_date: e.target.value || undefined });
                }}
                className="w-full rounded-lg border border-slate-800 bg-slate-950/80 px-2.5 py-1 text-xs text-slate-200 focus:border-indigo-500 focus:outline-none"
              />
            </div>

            {actionType === 'outlook_event' && (
              <div>
                <label className="block text-[11px] font-medium text-slate-400 mb-1 flex items-center gap-1">
                  <Clock className="h-3 w-3" /> Start Time
                </label>
                <input
                  type="time"
                  value={startTime}
                  disabled={isCompleted}
                  onChange={(e) => {
                    setStartTime(e.target.value);
                    handleSaveField({ start_time: e.target.value });
                  }}
                  className="w-full rounded-lg border border-slate-800 bg-slate-950/80 px-2.5 py-1 text-xs text-slate-200 focus:border-indigo-500 focus:outline-none"
                />
              </div>
            )}

            <div>
              <label className="block text-[11px] font-medium text-slate-400 mb-1 flex items-center gap-1">
                <Flag className="h-3 w-3" /> Priority
              </label>
              <select
                value={priority}
                disabled={isCompleted}
                onChange={(e) => {
                  const p = e.target.value as ActionPriority;
                  setPriority(p);
                  handleSaveField({ priority: p });
                }}
                className="w-full rounded-lg border border-slate-800 bg-slate-950/80 px-2.5 py-1 text-xs text-slate-200 focus:border-indigo-500 focus:outline-none"
              >
                <option value="low">Low</option>
                <option value="normal">Normal</option>
                <option value="high">High</option>
                <option value="urgent">Urgent</option>
              </select>
            </div>
          </div>
        )}

        {/* Error Display if Failed */}
        {action.error_message && (
          <div className="rounded-lg bg-rose-500/10 border border-rose-500/30 px-3 py-2 text-xs text-rose-300">
            {action.error_message}
          </div>
        )}

        {/* Footer Action Bar */}
        <div className="flex items-center justify-between pt-2 border-t border-slate-800/80">
          <div className="text-[11px] text-slate-500">
            {action.executed_at
              ? `Executed ${new Date(action.executed_at).toLocaleTimeString([], {
                  hour: '2-digit',
                  minute: '2-digit',
                })}`
              : `Created ${new Date(action.created_at).toLocaleTimeString([], {
                  hour: '2-digit',
                  minute: '2-digit',
                })}`}
          </div>

          <div className="flex items-center gap-2">
            {action.external_url && (
              <a
                href={action.external_url}
                target="_blank"
                rel="noopener noreferrer"
                className="inline-flex items-center gap-1 rounded-lg border border-slate-700 bg-slate-800 hover:bg-slate-700 px-3 py-1.5 text-xs font-medium text-slate-200 transition"
              >
                Open in App <ExternalLink className="h-3 w-3" />
              </a>
            )}

            {!isCompleted && (
              <button
                type="button"
                disabled={isExecuting}
                onClick={() => onExecute(buildCurrentAction())}
                className={`inline-flex items-center gap-1.5 rounded-lg px-3.5 py-1.5 text-xs font-semibold shadow-sm transition disabled:opacity-50 ${meta.buttonClass}`}
              >
                {isExecuting ? (
                  <>
                    <Loader2 className="h-3.5 w-3.5 animate-spin" />
                    Processing...
                  </>
                ) : (
                  <>
                    <Send className="h-3.5 w-3.5" />
                    {meta.buttonLabel}
                  </>
                )}
              </button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
