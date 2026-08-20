'use client';

import { useCallback, useEffect, useState, type FormEvent } from 'react';
import { useRouter } from 'next/navigation';
import { getContainer } from '@/infrastructure/container';
import { downloadWorkbook } from '@/infrastructure/download';
import {
  leaderboardToViews,
  type LeaderboardMemberView,
} from '@/interface-adapters/presenters/leaderboard.presenter';
import { BelloLogo } from '@/components/bello-logo';
import { GradeBadge } from '@/components/grade-badge';

function rankClass(rank: number): string {
  if (rank === 1) return 'bg-amber-400/20 text-amber-300 ring-amber-400/40';
  if (rank === 2) return 'bg-slate-300/15 text-slate-200 ring-slate-300/40';
  if (rank === 3) return 'bg-orange-400/15 text-orange-300 ring-orange-400/40';
  return 'bg-white/5 text-slate-400 ring-white/10';
}

const tableHead =
  'px-3 py-2.5 text-left text-[11px] font-semibold uppercase tracking-wider text-slate-400 whitespace-nowrap';

export default function HomePage() {
  const router = useRouter();
  const [members, setMembers] = useState<LeaderboardMemberView[]>([]);
  const [name, setName] = useState('');
  const [loading, setLoading] = useState(true);
  const [submitting, setSubmitting] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const load = useCallback(async () => {
    const entries = await getContainer().calculateLeaderboard.execute();
    setMembers(leaderboardToViews(entries));
    setLoading(false);
  }, []);

  useEffect(() => {
    let active = true;
    getContainer()
      .calculateLeaderboard.execute()
      .then((entries) => {
        if (active) setMembers(leaderboardToViews(entries));
      })
      .catch(() => {
        /* keep current list */
      })
      .finally(() => {
        if (active) setLoading(false);
      });
    return () => {
      active = false;
    };
  }, []);

  async function addMember(e: FormEvent): Promise<void> {
    e.preventDefault();
    if (!name.trim()) return;
    setSubmitting(true);
    setError(null);
    try {
      await getContainer().addMember.execute({ name: name.trim() });
      setName('');
      await load();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to add member');
    } finally {
      setSubmitting(false);
    }
  }

  async function handleExport(): Promise<void> {
    setExporting(true);
    setError(null);
    try {
      const bytes = await getContainer().exportEvaluationToExcel.execute();
      downloadWorkbook(bytes);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Export failed. Check that the template is bundled.');
    } finally {
      setExporting(false);
    }
  }

  return (
    <main className="animate-bello-in mx-auto w-full max-w-6xl px-4 pb-24 pt-10 sm:px-6">
      <header className="mb-8 flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-3">
          <BelloLogo size={46} />
          <div>
            <h1 className="text-3xl font-black tracking-tight text-white">Bello</h1>
            <p className="text-sm text-slate-400">Member Evaluation &amp; Tracking</p>
          </div>
        </div>
        <button
          type="button"
          onClick={() => void handleExport()}
          disabled={exporting}
          className="rounded-xl bg-gradient-to-r from-indigo-500 to-violet-500 px-4 py-2 text-sm font-semibold text-white shadow-lg shadow-indigo-950/40 transition-all hover:opacity-90 disabled:cursor-not-allowed disabled:opacity-40"
        >
          {exporting ? 'Exporting…' : 'Export Excel'}
        </button>
      </header>

      <section className="rounded-2xl border border-white/10 bg-white/[0.04] p-5 backdrop-blur">
        <h2 className="mb-3 text-lg font-bold text-white">Add member</h2>
        <form className="flex flex-wrap items-center gap-2" onSubmit={addMember}>
          <input
            className="min-w-48 flex-1 rounded-xl border border-white/10 bg-slate-950/60 px-3 py-2 text-sm text-slate-100 outline-none transition-colors placeholder:text-slate-500 focus:border-violet-400/60"
            placeholder="Member name"
            value={name}
            maxLength={100}
            required
            onChange={(e) => setName(e.target.value)}
          />
          <button
            type="submit"
            disabled={submitting || !name.trim()}
            className="rounded-xl bg-gradient-to-r from-indigo-500 to-violet-500 px-5 py-2 text-sm font-semibold text-white shadow-lg shadow-indigo-950/40 transition-all hover:opacity-90 disabled:cursor-not-allowed disabled:opacity-40"
          >
            {submitting ? 'Adding…' : 'Add'}
          </button>
        </form>
        {error && <p className="mt-3 text-sm text-rose-300">{error}</p>}
      </section>

      <section className="mt-4 rounded-2xl border border-white/10 bg-white/[0.04] p-5 backdrop-blur">
        <h2 className="mb-3 text-lg font-bold text-white">Leaderboard</h2>
        {loading ? (
          <div className="py-16 text-center text-sm text-slate-400">Loading…</div>
        ) : members.length === 0 ? (
          <div className="rounded-xl border border-dashed border-white/10 px-3 py-10 text-center text-sm text-slate-500">
            No members yet. Add the first member above.
          </div>
        ) : (
          <div className="-mx-5 overflow-x-auto px-5">
            <table className="w-full min-w-200 border-collapse text-sm">
              <thead>
                <tr className="border-b border-white/10">
                  <th className={tableHead}>#</th>
                  <th className={tableHead}>Member</th>
                  <th className={tableHead}>Tech /50</th>
                  <th className={tableHead}>Visits /20</th>
                  <th className={tableHead}>Meetings /10</th>
                  <th className={tableHead}>Interact /10</th>
                  <th className={tableHead}>Respect /10</th>
                  <th className={tableHead}>Bonus /10</th>
                  <th className={tableHead}>Total /110</th>
                  <th className={tableHead}>%</th>
                  <th className={tableHead}>Grade</th>
                </tr>
              </thead>
              <tbody>
                {members.map((member) => (
                  <tr
                    key={member.id}
                    onClick={() => router.push(`/members/${member.id}`)}
                    className="cursor-pointer border-b border-white/5 transition-colors last:border-0 hover:bg-white/[0.05]"
                  >
                    <td className="px-3 py-2.5 whitespace-nowrap">
                      <span
                        className={`inline-flex h-7 w-7 items-center justify-center rounded-full text-xs font-bold ring-1 ${rankClass(member.rank)}`}
                      >
                        {member.rank}
                      </span>
                    </td>
                    <td className="px-3 py-2.5 font-semibold text-violet-300">{member.name}</td>
                    <td className="px-3 py-2.5 whitespace-nowrap text-slate-300">{member.technical}</td>
                    <td className="px-3 py-2.5 whitespace-nowrap text-slate-300">{member.fieldVisits}</td>
                    <td className="px-3 py-2.5 whitespace-nowrap text-slate-300">{member.meetings}</td>
                    <td className="px-3 py-2.5 whitespace-nowrap text-slate-300">{member.interaction}</td>
                    <td className="px-3 py-2.5 whitespace-nowrap text-slate-300">{member.respectHierarchy}</td>
                    <td className="px-3 py-2.5 whitespace-nowrap text-slate-300">{member.bonus}</td>
                    <td className="px-3 py-2.5 text-[15px] font-extrabold whitespace-nowrap text-white">
                      {member.total}
                    </td>
                    <td className="px-3 py-2.5 whitespace-nowrap text-slate-300">{member.percentage}%</td>
                    <td className="px-3 py-2.5 whitespace-nowrap">
                      <GradeBadge grade={member.grade} />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </section>

      <footer className="mt-10 flex items-center justify-center gap-2 text-xs text-slate-500">
        <BelloLogo size={18} />
        <span>Bello &middot; offline-first member evaluation &middot; scores out of 110</span>
      </footer>
    </main>
  );
}