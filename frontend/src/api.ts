const BASE = ''

export interface UploadResult {
  file_id: string
  filename: string
  size: number
}

export interface ModeResult {
  mode: string
  count: number
  sheet: string
}

export interface JobResult {
  modes: ModeResult[]
  total: number
  output_file: string
}

export interface Job {
  job_id: string
  status: 'pending' | 'running' | 'completed' | 'failed'
  progress: string
  created_at: string
  completed_at: string | null
  error: string | null
  result: JobResult | null
}

export interface RunOptions {
  file_id: string
  mode: 'e2e_max' | 'e2e_short' | 'output_short' | 'multi_responses' | 'all'
  sheet?: string
  root?: string
  max_depth: number
  gen_data: boolean
}

async function _json<T>(res: Response): Promise<T> {
  if (!res.ok) {
    const text = await res.text()
    // Vite proxy returns 500 with empty body when it cannot reach the backend
    if (!text.trim() && res.status === 500) {
      throw new Error('Cannot reach the backend server. Make sure it is running:\n  python3 -m uvicorn backend.app:app --reload --port 8000')
    }
    let detail = text
    try {
      detail = JSON.parse(text)?.detail ?? text
    } catch {}
    throw new Error(detail || `HTTP ${res.status}`)
  }
  return res.json() as Promise<T>
}

export async function uploadFile(file: File): Promise<UploadResult> {
  const fd = new FormData()
  fd.append('file', file)
  const res = await fetch(`${BASE}/api/upload`, { method: 'POST', body: fd })
  return _json<UploadResult>(res)
}

export async function startPipeline(opts: RunOptions): Promise<{ job_id: string }> {
  const res = await fetch(`${BASE}/api/run`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(opts),
  })
  return _json<{ job_id: string }>(res)
}

export async function getJob(jobId: string): Promise<Job> {
  const res = await fetch(`${BASE}/api/jobs/${jobId}`)
  return _json<Job>(res)
}

export function downloadUrl(jobId: string): string {
  return `${BASE}/api/jobs/${jobId}/download`
}
