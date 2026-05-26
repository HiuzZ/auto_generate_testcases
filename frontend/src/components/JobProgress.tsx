import { useEffect, useRef, useState } from 'react'
import { getJob, Job } from '../api'

interface Props {
  jobId: string
  onDone: () => void
  onFailed: (err: string) => void
}

const POLL_MS = 1500

export default function JobProgress({ jobId, onDone, onFailed }: Props) {
  const [job, setJob] = useState<Job | null>(null)
  const [fetchError, setFetchError] = useState<string | null>(null)
  const timerRef = useRef<ReturnType<typeof setTimeout> | null>(null)

  useEffect(() => {
    let cancelled = false

    async function poll() {
      try {
        const data = await getJob(jobId)
        if (cancelled) return
        setJob(data)
        setFetchError(null)

        if (data.status === 'completed') {
          onDone()
          return
        }
        if (data.status === 'failed') {
          onFailed(data.error ?? 'Unknown error')
          return
        }
        timerRef.current = setTimeout(poll, POLL_MS)
      } catch (err: unknown) {
        if (cancelled) return
        setFetchError(err instanceof Error ? err.message : 'Failed to fetch job status')
        timerRef.current = setTimeout(poll, POLL_MS * 2)
      }
    }

    poll()
    return () => {
      cancelled = true
      if (timerRef.current) clearTimeout(timerRef.current)
    }
  }, [jobId, onDone, onFailed])

  const status = job?.status ?? 'pending'
  const progress = job?.progress ?? 'Waiting for worker…'

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 28 }}>
        <h2 style={{ fontSize: 22, fontWeight: 700 }}>Generating Test Cases</h2>
        <StatusBadge status={status} />
      </div>

      {/* Animated progress area */}
      <div style={{
        background: 'rgba(0,122,255,0.06)',
        border: '1px solid rgba(0,122,255,0.12)',
        borderRadius: 14,
        padding: '28px 24px',
        textAlign: 'center',
        marginBottom: 24,
      }}>
        {status === 'running' || status === 'pending' ? (
          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 16 }}>
            <div className="spinner" style={{ width: 40, height: 40, borderWidth: 4 }} />
            <p style={{ fontSize: 16, color: '#3C3C43', fontWeight: 500 }}>{progress}</p>
          </div>
        ) : (
          <p style={{ fontSize: 16, color: '#3C3C43' }}>{progress}</p>
        )}
      </div>

      {/* Job metadata */}
      <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap' }}>
        <MetaRow label="Job ID" value={jobId.slice(0, 8) + '…'} />
        {job?.created_at && (
          <MetaRow label="Started" value={new Date(job.created_at + 'Z').toLocaleTimeString()} />
        )}
      </div>

      {fetchError && (
        <div style={{
          marginTop: 16,
          padding: '10px 14px',
          background: 'rgba(255,149,0,0.1)',
          border: '1px solid rgba(255,149,0,0.2)',
          borderRadius: 10,
          color: '#FF9500',
          fontSize: 13,
        }}>
          ⚠️ {fetchError} — retrying…
        </div>
      )}
    </div>
  )
}

function StatusBadge({ status }: { status: Job['status'] }) {
  const map = {
    pending:   { cls: 'badge-pending',   icon: '…',  label: 'Pending' },
    running:   { cls: 'badge-running',   icon: null, label: 'Running' },
    completed: { cls: 'badge-completed', icon: '✓',  label: 'Done' },
    failed:    { cls: 'badge-failed',    icon: '✕',  label: 'Failed' },
  }
  const { cls, icon, label } = map[status]
  return (
    <span className={`badge ${cls}`}>
      {icon === null ? <span className="dot-pulse" /> : icon}
      {label}
    </span>
  )
}

function MetaRow({ label, value }: { label: string; value: string }) {
  return (
    <div style={{ display: 'flex', gap: 6, alignItems: 'baseline' }}>
      <span style={{ fontSize: 12, color: '#8E8E93', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.04em' }}>
        {label}
      </span>
      <span style={{ fontSize: 14, color: '#3C3C43', fontFamily: 'monospace' }}>{value}</span>
    </div>
  )
}
