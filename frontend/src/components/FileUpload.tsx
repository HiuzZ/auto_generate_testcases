import { useRef, useState, DragEvent } from 'react'
import { uploadFile, UploadResult } from '../api'

interface Props {
  onUploaded: (info: UploadResult) => void
}

export default function FileUpload({ onUploaded }: Props) {
  const inputRef = useRef<HTMLInputElement>(null)
  const [dragging, setDragging] = useState(false)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)

  async function handleFile(file: File) {
    setError(null)
    setLoading(true)
    try {
      const result = await uploadFile(file)
      onUploaded(result)
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : 'Upload failed')
    } finally {
      setLoading(false)
    }
  }

  function onInputChange(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (file) handleFile(file)
  }

  function onDrop(e: DragEvent) {
    e.preventDefault()
    setDragging(false)
    const file = e.dataTransfer.files?.[0]
    if (file) handleFile(file)
  }

  return (
    <div>
      <h2 style={{ fontSize: 22, fontWeight: 700, marginBottom: 6 }}>Upload Excel Script</h2>
      <p style={{ fontSize: 15, color: '#8E8E93', marginBottom: 24 }}>
        Drop your voicebot dialog Excel file (.xlsx) here.
      </p>

      <div
        onClick={() => !loading && inputRef.current?.click()}
        onDragOver={(e) => { e.preventDefault(); setDragging(true) }}
        onDragLeave={() => setDragging(false)}
        onDrop={onDrop}
        style={{
          border: `2px dashed ${dragging ? '#007AFF' : 'rgba(0,0,0,0.15)'}`,
          borderRadius: 16,
          padding: '48px 32px',
          textAlign: 'center',
          cursor: loading ? 'default' : 'pointer',
          background: dragging ? 'rgba(0,122,255,0.05)' : 'rgba(255,255,255,0.4)',
          transition: 'all 0.2s',
        }}
      >
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx,.xlsm,.xls"
          style={{ display: 'none' }}
          onChange={onInputChange}
        />

        {loading ? (
          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 12 }}>
            <div className="spinner" />
            <p style={{ fontSize: 15, color: '#8E8E93' }}>Uploading…</p>
          </div>
        ) : (
          <>
            <div style={{ fontSize: 48, marginBottom: 12 }}>📄</div>
            <p style={{ fontSize: 17, fontWeight: 600, color: dragging ? '#007AFF' : '#3C3C43' }}>
              {dragging ? 'Drop to upload' : 'Drag & drop or click to browse'}
            </p>
            <p style={{ fontSize: 13, color: '#8E8E93', marginTop: 6 }}>
              Supports .xlsx, .xlsm, .xls
            </p>
          </>
        )}
      </div>

      {error && (
        <div style={{
          marginTop: 16,
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
    </div>
  )
}
