'use client';

import { useState, type FormEvent, type ReactNode } from 'react';

export interface EntryItem {
  id: string;
  name?: string;
  date: string;
  score: number;
}

export interface EntryFormInput {
  name?: string;
  date: string;
  score: number;
}

const inputBase =
  'rounded-xl border border-white/10 bg-slate-950/60 px-3 py-2 text-sm text-slate-100 outline-none transition-colors placeholder:text-slate-500 focus:border-violet-400/60 disabled:cursor-not-allowed disabled:opacity-50';

/**
 * Field Visits (kind="visit") and Meetings (kind="meeting") share one editable
 * list. A visit collects a Location/Event name + a date; a meeting only a date.
 * The date picker is a native <input type="date"> so it works offline.
 */
export function EntrySection({
  title,
  description,
  kind,
  entries,
  max,
  disabled,
  onAdd,
  onUpdate,
  onRemove,
}: {
  title: string;
  description: ReactNode;
  kind: 'visit' | 'meeting';
  entries: EntryItem[];
  max?: number;
  disabled?: boolean;
  onAdd: (input: EntryFormInput) => Promise<unknown>;
  onUpdate: (id: string, input: EntryFormInput) => Promise<unknown>;
  onRemove: (id: string) => Promise<unknown>;
}) {
  const isVisit = kind === 'visit';
  // 1 = complete/attended, 0 = incomplete/absent (values are NOT inverted).
  const scoreOptions = isVisit
    ? [
        { value: 1, label: '1 - Complete' },
        { value: 0, label: '0 - Incomplete' },
      ]
    : [
        { value: 1, label: '1 - Attended' },
        { value: 0, label: '0 - Absent' },
      ];
  const statusLabel = (score: number): string =>
    isVisit
      ? score === 1
        ? 'Complete'
        : 'Incomplete'
      : score === 1
        ? 'Attended'
        : 'Absent';
  const emptyForm = (): EntryFormInput => ({ name: '', date: '', score: 1 });

  const [add, setAdd] = useState<EntryFormInput>(emptyForm);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [edit, setEdit] = useState<EntryFormInput>(emptyForm);
  const [busy, setBusy] = useState(false);
  const atLimit = max != null && entries.length >= max;

  function setAddField<K extends keyof EntryFormInput>(key: K, value: EntryFormInput[K]): void {
    setAdd((f) => ({ ...f, [key]: value }));
  }
  function setEditField<K extends keyof EntryFormInput>(key: K, value: EntryFormInput[K]): void {
    setEdit((f) => ({ ...f, [key]: value }));
  }

  async function handleAdd(e: FormEvent): Promise<void> {
    e.preventDefault();
    if (atLimit || !add.date || (isVisit && !add.name?.trim())) return;
    setBusy(true);
    try {
      await onAdd({ ...add, name: add.name?.trim() || undefined, score: Number(add.score) });
      setAdd(emptyForm());
    } finally {
      setBusy(false);
    }
  }

  function startEdit(item: EntryItem): void {
    setEditingId(item.id);
    setEdit({ name: item.name ?? '', date: item.date, score: item.score });
  }

  async function handleEdit(e: FormEvent): Promise<void> {
    e.preventDefault();
    if (!editingId || !edit.date || (isVisit && !edit.name?.trim())) return;
    setBusy(true);
    try {
      await onUpdate(editingId, { ...edit, name: edit.name?.trim() || undefined, score: Number(edit.score) });
      setEditingId(null);
    } finally {
      setBusy(false);
    }
  }

  const dateField = (
    value: string,
    onChange: (value: string) => void,
    required = true,
  ) => (
    <input
      type="date"
      className={inputBase}
      value={value}
      required={required}
      disabled={disabled || busy}
      onChange={(e) => onChange(e.target.value)}
      aria-label="Date"
    />
  );

  const scoreSelect = (
    value: number,
    onChange: (value: number) => void,
    disabledNow = false,
  ) => (
    <select
      className={`${inputBase} cursor-pointer`}
      value={value}
      disabled={disabled || busy || disabledNow}
      onChange={(e) => onChange(Number(e.target.value))}
      aria-label="Score"
    >
      {scoreOptions.map((option) => (
        <option key={option.value} value={option.value}>
          {option.label}
        </option>
      ))}
    </select>
  );

  return (
    <section className="rounded-2xl border border-white/10 bg-white/[0.04] p-5 backdrop-blur">
      <h2 className="mb-1 text-lg font-bold text-white">{title}</h2>
      <p className="mb-4 text-sm text-slate-400">{description}</p>

      <form className="flex flex-wrap items-center gap-2" onSubmit={handleAdd}>
        {isVisit && (
          <input
            className={`${inputBase} min-w-44 flex-1`}
            placeholder="Location / event name"
            value={add.name ?? ''}
            maxLength={200}
            required
            disabled={disabled || busy || atLimit}
            onChange={(e) => setAddField('name', e.target.value)}
          />
        )}
        {dateField(add.date, (v) => setAddField('date', v), false)}
        {scoreSelect(add.score, (v) => setAddField('score', v), false)}
        <button
          type="submit"
          disabled={disabled || busy || atLimit}
          className="rounded-xl bg-gradient-to-r from-indigo-500 to-violet-500 px-4 py-2 text-sm font-semibold text-white shadow-lg shadow-indigo-950/40 transition-all hover:opacity-90 disabled:cursor-not-allowed disabled:opacity-40"
        >
          {busy ? 'Saving…' : isVisit ? 'Add visit' : 'Add meeting'}
        </button>
      </form>

      {atLimit && (
        <p className="mt-2 text-xs text-amber-300">Maximum {max} {isVisit ? 'visits' : 'meetings'} reached.</p>
      )}

      <ul className="mt-4 space-y-2">
        {entries.length === 0 ? (
          <li className="rounded-xl border border-dashed border-white/10 px-3 py-4 text-center text-sm text-slate-500">
            No {isVisit ? 'visits' : 'meetings'} recorded yet.
          </li>
        ) : (
          entries.map((item) =>
            editingId === item.id ? (
              <li key={item.id}>
                <form className="flex flex-wrap items-center gap-2" onSubmit={handleEdit}>
                  {isVisit && (
                    <input
                      className={`${inputBase} min-w-44 flex-1`}
                      value={edit.name ?? ''}
                      maxLength={200}
                      required
                      disabled={disabled || busy}
                      onChange={(e) => setEditField('name', e.target.value)}
                    />
                  )}
                  {dateField(edit.date, (v) => setEditField('date', v))}
                  {scoreSelect(edit.score, (v) => setEditField('score', v))}
                  <button
                    type="submit"
                    disabled={disabled || busy}
                    className="rounded-xl bg-gradient-to-r from-emerald-500 to-teal-500 px-4 py-2 text-sm font-semibold text-white shadow-lg shadow-emerald-950/40 transition-all hover:opacity-90 disabled:cursor-not-allowed disabled:opacity-40"
                  >
                    Save
                  </button>
                  <button
                    type="button"
                    disabled={disabled || busy}
                    onClick={() => setEditingId(null)}
                    className="rounded-xl border border-white/10 px-4 py-2 text-sm font-semibold text-slate-300 transition-colors hover:bg-white/5 disabled:cursor-not-allowed disabled:opacity-40"
                  >
                    Cancel
                  </button>
                </form>
              </li>
            ) : (
              <li
                key={item.id}
                className="group flex flex-wrap items-center gap-3 rounded-xl border border-white/5 bg-slate-950/40 px-3 py-2.5 transition-colors hover:border-white/10"
              >
                <div className="min-w-0 flex-1">
                  {isVisit && item.name && (
                    <div className="truncate text-sm font-medium text-slate-100">{item.name}</div>
                  )}
                  <div className="text-xs text-slate-400">{item.date}</div>
                </div>
                <span
                  className={`rounded-lg px-2 py-0.5 text-xs font-bold ring-1 ${
                    item.score === 1
                      ? 'bg-emerald-500/15 text-emerald-300 ring-emerald-400/30'
                      : 'bg-rose-500/15 text-rose-300 ring-rose-400/30'
                  }`}
                >
                  {statusLabel(item.score)}
                </span>
                <button
                  type="button"
                  title="Edit"
                  disabled={disabled || busy}
                  onClick={() => startEdit(item)}
                  className="rounded-lg px-2 py-1 text-xs font-semibold text-slate-300 transition-all hover:bg-white/10 disabled:cursor-not-allowed disabled:opacity-40 sm:opacity-0 sm:group-hover:opacity-100"
                >
                  Edit
                </button>
                <button
                  type="button"
                  title="Remove"
                  disabled={disabled || busy}
                  onClick={() => void onRemove(item.id)}
                  className="rounded-lg px-2 py-1 text-xs font-semibold text-rose-300 transition-all hover:bg-rose-500/15 disabled:cursor-not-allowed disabled:opacity-40 sm:opacity-0 sm:group-hover:opacity-100"
                >
                  Remove
                </button>
              </li>
            ),
          )
        )}
      </ul>
    </section>
  );
}