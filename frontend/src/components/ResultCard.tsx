import { useEffect, useState } from 'react'
import { getJob, downloadUrl, Job } from '../api'

interface Props {
  jobId: string
  onReset: () => void
}

export default function ResultCard({ jobId, onReset }: Props) {
  const [job, setJob] = useState<Job | null>(null)

  useEffect(() => {
    getJob(jobId).then(setJob).catch(() => {})
  }, [jobId])

  if (!job?.result) {
    return (
      <div style={{ textAlign: 'center', padding: 40 }}>
        <div className="spinner" />
      </div>
    )
  }

  const { result } = job
  const completedTime = job.completed_at
    ? new Date(job.completed_at + 'Z').toLocaleString()
    : ''

  return (
    <div>
      {/* Success header */}
      <div style={{ textAlign: 'center', marginBottom: 32 }}>
        <div style={{ fontSize: 56, marginBottom: 12 }}>🎉</div>
        <h2 style={{ fontSize: 26, fontWeight: 700, marginBottom: 4 }}>Pipeline Complete</h2>
        {completedTime && (
          <p style={{ fontSize: 13, color: '#8E8E93' }}>Finished at {completedTime}</p>
        )}
      </div>

      {/* Total summary */}
      <div style={{
        background: 'linear-gradient(135deg, rgba(0,122,255,0.1) 0%, rgba(52,199,89,0.08) 100%)',
        border: '1px solid rgba(0,122,255,0.15)',
        borderRadius: 16,
        padding: '20px 24px',
        textAlign: 'center',
        marginBottom: 20,
      }}>
        <div style={{ fontSize: 52, fontWeight: 800, color: '#007AFF', lineHeight: 1 }}>
          {result.total}
        </div>
        <div style={{ fontSize: 15, color: '#3C3C43', marginTop: 6, fontWeight: 500 }}>
          Total Test Cases Generated
        </div>
      </div>

      {/* Per-mode breakdown */}
      <div className="stat-grid" style={{ marginBottom: 28 }}>
        {result.modes.map((m) => (
          <div key={m.mode} className="stat-card">
            <div className="stat-value">{m.count}</div>
            <div className="stat-label">{modeLabel(m.mode)}</div>
          </div>
        ))}
      </div>

      {/* Download button */}
      <a
        href={downloadUrl(jobId)}
        download={result.output_file}
        className="btn btn-primary"
        style={{ width: '100%', fontSize: 18, padding: '16px', display: 'flex', marginBottom: 12 }}
      >
        ⬇  Download Excel
      </a>

      <button
        className="btn btn-ghost"
        style={{ width: '100%' }}
        onClick={onReset}
      >
        Generate Another File
      </button>
    </div>
  )
}

function modeLabel(mode: string): string {
  const map: Record<string, string> = {
    e2e_short: 'E2E Short',
    output_short: 'Output Short',
    multi_responses: 'Multi-response',
  }
  return map[mode] ?? mode
}
