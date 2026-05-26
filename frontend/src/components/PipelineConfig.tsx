import { useState } from 'react'
import { startPipeline, RunOptions } from '../api'

interface Props {
  fileId: string
  onJobStarted: (jobId: string) => void
}

type Mode = RunOptions['mode']

const MODES: { value: Mode; label: string; desc: string }[] = [
  { value: 'all', label: 'All (Recommended)', desc: 'E2E + Output + Multi-response' },
  { value: 'e2e_max', label: 'E2E Max', desc: 'End-to-end max cases without repeat-chain grouping' },
  { value: 'e2e_short', label: 'E2E Short', desc: 'End-to-end test cases' },
  { value: 'output_short', label: 'Output Short', desc: 'Bot output test cases' },
  { value: 'multi_responses', label: 'Multi-response', desc: 'Multi-turn dialog cases' },
]

export default function PipelineConfig({ fileId, onJobStarted }: Props) {
  const [mode, setMode] = useState<Mode>('all')
  const [sheet, setSheet] = useState('')
  const [root, setRoot] = useState('')
  const [maxDepth, setMaxDepth] = useState(200)
  const [genData, setGenData] = useState(false)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)

  async function handleRun() {
    setError(null)
    setLoading(true)
    try {
      const { job_id } = await startPipeline({
        file_id: fileId,
        mode,
        sheet: sheet.trim() || undefined,
        root: root.trim() || undefined,
        max_depth: maxDepth,
        gen_data: genData,
      })
      onJobStarted(job_id)
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : 'Failed to start pipeline')
      setLoading(false)
    }
  }

  return (
    <div>
      <h2 style={{ fontSize: 22, fontWeight: 700, marginBottom: 6 }}>Configure Pipeline</h2>
      <p style={{ fontSize: 15, color: '#8E8E93', marginBottom: 24 }}>
        Choose mode and optional overrides. Blanks auto-detect.
      </p>

      {/* Mode selector */}
      <div className="form-group" style={{ marginBottom: 24 }}>
        <label className="form-label">Pipeline Mode</label>
        <div className="mode-chips">
          {MODES.map((m) => (
            <button
              key={m.value}
              className={`mode-chip${mode === m.value ? ' active' : ''}`}
              onClick={() => setMode(m.value)}
              title={m.desc}
            >
              {m.label}
            </button>
          ))}
        </div>
      </div>

      {/* Advanced options */}
      <details style={{ marginBottom: 24 }}>
        <summary style={{
          cursor: 'pointer',
          fontSize: 15,
          fontWeight: 600,
          color: '#007AFF',
          listStyle: 'none',
          display: 'flex',
          alignItems: 'center',
          gap: 6,
          userSelect: 'none',
        }}>
          ⚙️ Advanced Options
        </summary>
        <div style={{ paddingTop: 20, display: 'flex', flexDirection: 'column', gap: 16 }}>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
            <div className="form-group">
              <label className="form-label">Sheet (optional)</label>
              <input
                className="form-input"
                placeholder="Auto-detect"
                value={sheet}
                onChange={(e) => setSheet(e.target.value)}
              />
            </div>
            <div className="form-group">
              <label className="form-label">Root Step (optional)</label>
              <input
                className="form-input"
                placeholder="Auto-detect"
                value={root}
                onChange={(e) => setRoot(e.target.value)}
              />
            </div>
          </div>

          <div className="form-group">
            <label className="form-label">Max Depth — {maxDepth}</label>
            <input
              type="range"
              min={10}
              max={500}
              step={10}
              value={maxDepth}
              onChange={(e) => setMaxDepth(Number(e.target.value))}
              style={{ width: '100%', accentColor: '#007AFF' }}
            />
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, color: '#8E8E93' }}>
              <span>10</span><span>500</span>
            </div>
          </div>
        </div>
      </details>

      {/* Generate test data toggle */}
      <div style={{
        background: 'rgba(255,255,255,0.5)',
        border: '1px solid rgba(0,0,0,0.07)',
        borderRadius: 14,
        padding: '16px 18px',
        marginBottom: 28,
      }}>
        <div className="toggle-row">
          <div className="toggle-info">
            <strong>Generate Test Data</strong>
            <span>Fill customer utterances via TLS Bot + LLM (slower)</span>
          </div>
          <label className="toggle">
            <input
              type="checkbox"
              checked={genData}
              onChange={(e) => setGenData(e.target.checked)}
            />
            <div className="toggle-track" />
            <div className="toggle-thumb" />
          </label>
        </div>
      </div>

      {error && (
        <div style={{
          marginBottom: 20,
          padding: '12px 16px',
          background: 'rgba(255,59,48,0.1)',
          border: '1px solid rgba(255,59,48,0.2)',
          borderRadius: 10,
          color: '#FF3B30',
          fontSize: 14,
        }}>
          {error}
        </div>
      )}

      <button
        className="btn btn-primary"
        style={{ width: '100%', fontSize: 18, padding: '16px' }}
        onClick={handleRun}
        disabled={loading}
      >
        {loading ? (
          <><div className="spinner" style={{ borderTopColor: '#fff', borderColor: 'rgba(255,255,255,0.3)' }} /> Starting…</>
        ) : (
          '▶  Run Pipeline'
        )}
      </button>
    </div>
  )
}
