import { useState, useCallback } from 'react'
import FileUpload from './components/FileUpload'
import PipelineConfig from './components/PipelineConfig'
import JobProgress from './components/JobProgress'
import ResultCard from './components/ResultCard'
import { UploadResult } from './api'

type Stage = 'upload' | 'configure' | 'running' | 'done' | 'error'

export default function App() {
  const [stage, setStage] = useState<Stage>('upload')
  const [fileInfo, setFileInfo] = useState<UploadResult | null>(null)
  const [jobId, setJobId] = useState<string | null>(null)
  const [globalError, setGlobalError] = useState<string | null>(null)

  const handleUploaded = useCallback((info: UploadResult) => {
    setFileInfo(info)
    setGlobalError(null)
    setStage('configure')
  }, [])

  const handleJobStarted = useCallback((id: string) => {
    setJobId(id)
    setStage('running')
  }, [])

  const handleJobDone = useCallback(() => {
    setStage('done')
  }, [])

  const handleJobFailed = useCallback((err: string) => {
    setGlobalError(err)
    setStage('error')
  }, [])

  const handleReset = useCallback(() => {
    setStage('upload')
    setFileInfo(null)
    setJobId(null)
    setGlobalError(null)
  }, [])

  return (
    <div style={{ minHeight: '100vh', padding: '40px 20px' }}>
      {/* Header */}
      <header style={{ textAlign: 'center', marginBottom: 48 }}>
        <div style={{ fontSize: 48, marginBottom: 12 }}>🧪</div>
        <h1 style={{ fontSize: 34, fontWeight: 700, letterSpacing: '-0.5px', color: '#000' }}>
          Test Case Generator
        </h1>
        <p style={{ fontSize: 17, color: '#3C3C43', marginTop: 8 }}>
          Upload an Excel script, run the pipeline, download your test cases.
        </p>
      </header>

      {/* Stepper */}
      <Stepper stage={stage} />

      {/* Main card */}
      <div style={{ maxWidth: 680, margin: '0 auto' }}>
        {(stage === 'upload') && (
          <div className="glass" style={{ padding: 32 }}>
            <FileUpload onUploaded={handleUploaded} />
          </div>
        )}

        {stage === 'configure' && fileInfo && (
          <div className="glass" style={{ padding: 32 }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 24 }}>
              <div>
                <p style={{ fontSize: 13, color: '#8E8E93', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.04em' }}>
                  File ready
                </p>
                <p style={{ fontSize: 17, fontWeight: 600, marginTop: 2 }}>{fileInfo.filename}</p>
                <p style={{ fontSize: 13, color: '#8E8E93' }}>{(fileInfo.size / 1024).toFixed(1)} KB</p>
              </div>
              <button className="btn btn-ghost btn-sm" onClick={handleReset}>
                Change file
              </button>
            </div>
            <div className="divider" />
            <PipelineConfig fileId={fileInfo.file_id} onJobStarted={handleJobStarted} />
          </div>
        )}

        {stage === 'running' && jobId && (
          <div className="glass" style={{ padding: 32 }}>
            <JobProgress
              jobId={jobId}
              onDone={handleJobDone}
              onFailed={handleJobFailed}
            />
          </div>
        )}

        {stage === 'done' && jobId && (
          <div className="glass" style={{ padding: 32 }}>
            <ResultCard jobId={jobId} onReset={handleReset} />
          </div>
        )}

        {stage === 'error' && (
          <div className="glass" style={{ padding: 32 }}>
            <div style={{ textAlign: 'center' }}>
              <div style={{ fontSize: 48, marginBottom: 16 }}>⚠️</div>
              <h2 style={{ fontSize: 22, fontWeight: 700, marginBottom: 8, color: '#FF3B30' }}>
                Pipeline Failed
              </h2>
              <p style={{ fontSize: 15, color: '#3C3C43', marginBottom: 24, whiteSpace: 'pre-wrap' }}>
                {globalError}
              </p>
              <button className="btn btn-primary" onClick={handleReset}>
                Start Over
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  )
}

function Stepper({ stage }: { stage: Stage }) {
  const steps = [
    { key: 'upload', label: 'Upload' },
    { key: 'configure', label: 'Configure' },
    { key: 'running', label: 'Generate' },
    { key: 'done', label: 'Download' },
  ]
  const order: Record<string, number> = { upload: 0, configure: 1, running: 2, done: 3, error: 3 }
  const current = order[stage] ?? 0

  return (
    <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', gap: 0, marginBottom: 40 }}>
      {steps.map((s, i) => {
        const done = i < current
        const active = i === current
        return (
          <div key={s.key} style={{ display: 'flex', alignItems: 'center' }}>
            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 6 }}>
              <div style={{
                width: 32,
                height: 32,
                borderRadius: '50%',
                background: done ? '#34C759' : active ? '#007AFF' : 'rgba(0,0,0,0.1)',
                color: done || active ? '#fff' : '#8E8E93',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                fontSize: 14,
                fontWeight: 700,
                transition: 'background 0.3s',
              }}>
                {done ? '✓' : i + 1}
              </div>
              <span style={{
                fontSize: 12,
                fontWeight: active ? 700 : 500,
                color: active ? '#007AFF' : done ? '#34C759' : '#8E8E93',
                transition: 'color 0.3s',
              }}>
                {s.label}
              </span>
            </div>
            {i < steps.length - 1 && (
              <div style={{
                width: 80,
                height: 2,
                background: done ? '#34C759' : 'rgba(0,0,0,0.1)',
                marginBottom: 18,
                transition: 'background 0.3s',
              }} />
            )}
          </div>
        )
      })}
    </div>
  )
}
