import type { SitpoJob, SitpoJobRequest } from '../types'

const API_BASE = (import.meta.env.VITE_SITPO_API_BASE || '/api').replace(/\/$/, '')

function fileUrl(path: string) {
  if (/^https?:\/\//.test(path)) return path
  return `${API_BASE}${path.startsWith('/') ? path : `/${path}`}`
}

async function readJson<T>(response: Response): Promise<T> {
  const data = await response.json().catch(() => ({}))

  if (!response.ok) {
    const message = typeof data?.detail === 'string' ? data.detail : `SITPO API 오류: ${response.status}`
    throw new Error(message)
  }

  return data as T
}

export async function createSitpoJob(request: SitpoJobRequest): Promise<SitpoJob> {
  const response = await fetch(`${API_BASE}/jobs`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(request),
  })

  return readJson<SitpoJob>(response)
}

export async function getSitpoJob(jobId: string): Promise<SitpoJob> {
  const response = await fetch(`${API_BASE}/jobs/${jobId}`)
  return readJson<SitpoJob>(response)
}

export function sitpoJobFileUrl(jobId: string, filename: string) {
  return fileUrl(`/jobs/${jobId}/files/${encodeURIComponent(filename)}`)
}

export function normalizeSitpoJob(job: SitpoJob): SitpoJob {
  return {
    ...job,
    files: job.files.map((file) => ({ ...file, url: fileUrl(file.url) })),
  }
}
