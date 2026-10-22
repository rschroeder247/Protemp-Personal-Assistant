'use client';

import React, { useState } from 'react';
import { Contact } from '@/lib/types';
import {
  Users,
  Plus,
  Trash2,
  X,
  Phone,
  Mail,
  Building2,
  RefreshCw,
  Search,
} from 'lucide-react';

interface ContactsModalProps {
  isOpen: boolean;
  onClose: () => void;
  contacts: Contact[];
  onAddOrUpdate: (contact: Partial<Contact> & { name: string }) => Promise<void>;
  onDelete: (id: string) => Promise<void>;
  onSyncM365Contacts: () => Promise<number>;
  isM365SignedIn: boolean;
}

export default function ContactsModal({
  isOpen,
  onClose,
  contacts,
  onAddOrUpdate,
  onDelete,
  onSyncM365Contacts,
  isM365SignedIn,
}: ContactsModalProps) {
  const [name, setName] = useState('');
  const [phone, setPhone] = useState('');
  const [email, setEmail] = useState('');
  const [role, setRole] = useState('');
  const [company, setCompany] = useState('Protemp');
  const [search, setSearch] = useState('');
  const [isSaving, setIsSaving] = useState(false);
  const [isSyncing, setIsSyncing] = useState(false);
  const [syncMessage, setSyncMessage] = useState<string | null>(null);

  if (!isOpen) return null;

  const handleCreate = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!name.trim()) return;
    setIsSaving(true);
    try {
      await onAddOrUpdate({
        name: name.trim(),
        phone: phone.trim() || undefined,
        email: email.trim() || undefined,
        role: role.trim() || undefined,
        company: company.trim() || 'Protemp',
      });
      setName('');
      setPhone('');
      setEmail('');
      setRole('');
    } finally {
      setIsSaving(false);
    }
  };

  const handleSync = async () => {
    setIsSyncing(true);
    setSyncMessage(null);
    try {
      const count = await onSyncM365Contacts();
      setSyncMessage(`Imported ${count} contact(s) from Microsoft 365.`);
    } catch (err: any) {
      setSyncMessage(err?.message || 'Could not import M365 contacts.');
    } finally {
      setIsSyncing(false);
    }
  };

  const filtered = contacts.filter(
    (c) =>
      c.name.toLowerCase().includes(search.toLowerCase()) ||
      (c.phone && c.phone.includes(search)) ||
      (c.email && c.email.toLowerCase().includes(search.toLowerCase()))
  );

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 backdrop-blur-sm p-4">
      <div className="flex max-h-[90vh] w-full max-w-3xl flex-col rounded-2xl border border-slate-800 bg-slate-900 shadow-2xl">
        {/* Header */}
        <div className="flex items-center justify-between border-b border-slate-800 px-6 py-4">
          <div className="flex items-center gap-2.5">
            <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-emerald-500/20 text-emerald-400">
              <Users className="h-5 w-5" />
            </div>
            <div>
              <h3 className="text-base font-semibold text-white">
                Contacts Directory (Supabase Synced)
              </h3>
              <p className="text-xs text-slate-400">
                When you say &ldquo;Send a WhatsApp to John&rdquo;, your assistant automatically matches John&apos;s phone number &amp; email here.
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

        {/* Body */}
        <div className="flex-1 overflow-y-auto p-6 space-y-6">
          {/* Add Contact Form */}
          <form
            onSubmit={handleCreate}
            className="rounded-xl border border-slate-800 bg-slate-950/60 p-4 space-y-3"
          >
            <div className="flex items-center justify-between">
              <span className="text-xs font-semibold uppercase tracking-wider text-slate-400">
                Add New Contact
              </span>
              {isM365SignedIn && (
                <button
                  type="button"
                  onClick={handleSync}
                  disabled={isSyncing}
                  className="inline-flex items-center gap-1.5 rounded-lg border border-sky-500/30 bg-sky-500/10 hover:bg-sky-500/20 px-2.5 py-1 text-xs font-medium text-sky-300 transition"
                >
                  <RefreshCw
                    className={`h-3.5 w-3.5 ${isSyncing ? 'animate-spin' : ''}`}
                  />
                  Import from Microsoft 365
                </button>
              )}
            </div>

            {syncMessage && (
              <div className="text-xs text-sky-300 bg-sky-500/10 border border-sky-500/20 rounded-lg px-3 py-1.5">
                {syncMessage}
              </div>
            )}

            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-2.5">
              <input
                type="text"
                required
                placeholder="Full Name (e.g. John Smith)"
                value={name}
                onChange={(e) => setName(e.target.value)}
                className="rounded-lg border border-slate-800 bg-slate-900 px-3 py-2 text-xs text-white placeholder-slate-500 focus:border-emerald-500 focus:outline-none"
              />
              <input
                type="text"
                placeholder="WhatsApp (+2782...)"
                value={phone}
                onChange={(e) => setPhone(e.target.value)}
                className="rounded-lg border border-slate-800 bg-slate-900 px-3 py-2 text-xs text-white placeholder-slate-500 focus:border-emerald-500 focus:outline-none"
              />
              <input
                type="email"
                placeholder="Email (john@protemp.co.za)"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                className="rounded-lg border border-slate-800 bg-slate-900 px-3 py-2 text-xs text-white placeholder-slate-500 focus:border-emerald-500 focus:outline-none"
              />
              <div className="flex gap-2">
                <input
                  type="text"
                  placeholder="Role / Company"
                  value={role}
                  onChange={(e) => setRole(e.target.value)}
                  className="w-full rounded-lg border border-slate-800 bg-slate-900 px-3 py-2 text-xs text-white placeholder-slate-500 focus:border-emerald-500 focus:outline-none"
                />
                <button
                  type="submit"
                  disabled={isSaving || !name.trim()}
                  className="inline-flex shrink-0 items-center gap-1 rounded-lg bg-emerald-600 hover:bg-emerald-500 disabled:opacity-40 px-3.5 py-2 text-xs font-semibold text-white transition"
                >
                  <Plus className="h-4 w-4" /> Add
                </button>
              </div>
            </div>
          </form>

          {/* Search & List */}
          <div>
            <div className="flex items-center justify-between mb-3">
              <span className="text-xs font-medium text-slate-400">
                Saved Contacts ({filtered.length})
              </span>
              <div className="relative w-64">
                <Search className="h-3.5 w-3.5 text-slate-500 absolute left-3 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  placeholder="Search contacts..."
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  className="w-full rounded-lg border border-slate-800 bg-slate-950 pl-8 pr-3 py-1.5 text-xs text-white placeholder-slate-500 focus:border-indigo-500 focus:outline-none"
                />
              </div>
            </div>

            {filtered.length === 0 ? (
              <div className="rounded-xl border border-dashed border-slate-800 py-10 text-center text-xs text-slate-500">
                No contacts saved yet. Add a contact above so voice commands can automatically resolve phone numbers and emails.
              </div>
            ) : (
              <div className="divide-y divide-slate-800/80 rounded-xl border border-slate-800 bg-slate-950/40">
                {filtered.map((c) => (
                  <div
                    key={c.id}
                    className="flex flex-wrap items-center justify-between gap-3 px-4 py-3 hover:bg-slate-900/60"
                  >
                    <div>
                      <div className="text-sm font-medium text-white flex items-center gap-2">
                        {c.name}
                        {c.role && (
                          <span className="rounded bg-slate-800 px-2 py-0.5 text-[10px] text-slate-400">
                            {c.role}
                          </span>
                        )}
                      </div>
                      <div className="mt-1 flex flex-wrap items-center gap-4 text-xs text-slate-400">
                        {c.phone && (
                          <span className="inline-flex items-center gap-1 text-emerald-400">
                            <Phone className="h-3 w-3" /> {c.phone}
                          </span>
                        )}
                        {c.email && (
                          <span className="inline-flex items-center gap-1 text-sky-400">
                            <Mail className="h-3 w-3" /> {c.email}
                          </span>
                        )}
                        {c.company && (
                          <span className="inline-flex items-center gap-1 text-slate-500">
                            <Building2 className="h-3 w-3" /> {c.company}
                          </span>
                        )}
                      </div>
                    </div>

                    <button
                      type="button"
                      onClick={() => onDelete(c.id)}
                      className="rounded-lg p-1.5 text-slate-500 hover:bg-slate-800 hover:text-rose-400"
                      title="Delete Contact"
                    >
                      <Trash2 className="h-4 w-4" />
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
