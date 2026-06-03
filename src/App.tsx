import { useEffect, useMemo, useState } from 'react'
import './App.css'
import { getSubjectsForGrade, getUnitsForGradeSubject, gradeOptions } from './data/elementaryCurriculum'
import { workflowSteps } from './data/sampleProject'
import { downloadText, exportMarkdown } from './lib/exporters'
import { generatePptx } from './lib/pptxGenerator'
import { createSitpoJob, getSitpoJob, normalizeSitpoJob } from './lib/sitpoApi'
import type { SitpoJob, SitpoJobFile, SitpoProject, WorkflowStatus, WorkflowStep } from './types'

function statusFor(index: number, currentIndex: number): WorkflowStatus {
  if (index < currentIndex) return '완료'
  if (index === currentIndex) return '진행 중'
  return '대기'
}

function jobLabel(job?: SitpoJob | null) {
  if (!job) return '요청 대기'
  if (job.status === 'queued') return '대기 중'
  if (job.status === 'running') return 'Codex 작업 중'
  if (job.status === 'succeeded') return '완료'
  return '실패'
}

function latestLog(job?: SitpoJob | null) {
  return job?.logs.at(-1) ?? ''
}

function serverPptxFile(job?: SitpoJob | null) {
  return job?.files.find((file) => file.filename.endsWith('.pptx'))
}

function App() {
  const [project, setProject] = useState<SitpoProject | null>(null)
  const [job, setJob] = useState<SitpoJob | null>(null)
  const [currentView, setCurrentView] = useState<'input' | 'preview'>('input')
  const [grade, setGrade] = useState('초5')
  const [subject, setSubject] = useState('과학')
  const [unit, setUnit] = useState('생물과 환경')
  const [topic, setTopic] = useState('생태계의 구성 요소')
  const [slideCount, setSlideCount] = useState(9)
  const [style, setStyle] = useState('프리미엄 과학 탐험 노트형')
  const [statusMessage, setStatusMessage] = useState('')
  const [loading, setLoading] = useState(false)

  const subjectOptions = useMemo(() => getSubjectsForGrade(grade), [grade])
  const unitOptions = useMemo(() => getUnitsForGradeSubject(grade, subject), [grade, subject])
  const currentStep = project ? workflowSteps.find(step => step.id === project.currentStep) || workflowSteps[0] : workflowSteps[0]
  const currentStepIndex = project ? Math.max(0, workflowSteps.findIndex(step => step.id === project.currentStep)) : 0
  const progressPercent = job?.status === 'succeeded' || job?.status === 'failed'
    ? 100
    : Math.max(12, Math.round(((currentStepIndex + 1) / workflowSteps.length) * 100))

  useEffect(() => {
    if (!job || !['queued', 'running'].includes(job.status)) return undefined

    const interval = window.setInterval(async () => {
      try {
        const nextJob = normalizeSitpoJob(await getSitpoJob(job.id))
        setJob(nextJob)
        setProject(nextJob.result ?? null)
        setStatusMessage(latestLog(nextJob) || `${jobLabel(nextJob)}입니다. Codex 작업은 시간이 걸릴 수 있습니다.`)
      } catch (error) {
        console.error('SITPO job polling failed:', error)
        setStatusMessage('작업 상태를 가져오지 못했습니다. 서버 연결을 확인해 주세요.')
      }
    }, 4000)

    return () => window.clearInterval(interval)
  }, [job])

  const handleGradeChange = (nextGrade: string) => {
    const nextSubjects = getSubjectsForGrade(nextGrade)
    const nextSubject = nextSubjects[0]?.subject ?? ''
    const nextUnit = nextSubjects[0]?.units[0] ?? ''

    setGrade(nextGrade)
    setSubject(nextSubject)
    setUnit(nextUnit)
    setTopic(nextUnit)
  }

  const handleSubjectChange = (nextSubject: string) => {
    const nextUnit = getUnitsForGradeSubject(grade, nextSubject)[0] ?? ''

    setSubject(nextSubject)
    setUnit(nextUnit)
    setTopic(nextUnit)
  }

  const handleGeneratePlan = async () => {
    setLoading(true)
    setStatusMessage('서버에 Codex 생성 작업을 요청하고 있습니다...')

    try {
      const createdJob = normalizeSitpoJob(await createSitpoJob({ grade, subject, unit, topic, style, slideCount }))
      setJob(createdJob)
      setProject(createdJob.result ?? null)
      setCurrentView('preview')
      setStatusMessage('작업이 등록되었습니다. Codex 작업 중에는 몇 분 정도 걸릴 수 있습니다.')
    } catch (error) {
      console.error('SITPO job create failed:', error)
      setStatusMessage(error instanceof Error ? error.message : '생성 요청에 실패했습니다.')
    } finally {
      setLoading(false)
    }
  }

  const handleProceed = async () => {
    if (!job) return

    setLoading(true)
    setStatusMessage('서버에서 실제 작업 상태를 새로고침합니다...')
    try {
      const refreshedJob = normalizeSitpoJob(await getSitpoJob(job.id))
      setJob(refreshedJob)
      setProject(refreshedJob.result ?? null)
      setStatusMessage(latestLog(refreshedJob) || `${jobLabel(refreshedJob)}입니다.`)
    } catch (error) {
      console.error('SITPO job refresh failed:', error)
      setStatusMessage('작업 상태 새로고침에 실패했습니다.')
    } finally {
      setLoading(false)
    }
  }

  const handleGeneratePptx = async () => {
    const pptxFile = serverPptxFile(job)
    if (pptxFile) {
      window.location.href = pptxFile.url
      setStatusMessage('서버 PPTX 다운로드를 시작했습니다.')
      return
    }

    if (!project) {
      setStatusMessage('서버 PPTX가 아직 없습니다. Codex 작업이 끝난 뒤 다시 시도해 주세요.')
      return
    }

    setStatusMessage('서버 PPTX가 없어 브라우저 미리보기 PPTX를 생성합니다...')
    setLoading(true)
    try {
      const pptxBlob = await generatePptx(project)
      downloadText(`${project.id}.pptx`, pptxBlob, 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
      setStatusMessage('브라우저 미리보기 PPTX 다운로드를 시작했습니다.')
    } catch (error) {
      console.error('Error generating PPTX:', error)
      setStatusMessage('PPTX 생성 중 오류가 발생했습니다.')
    } finally {
      setLoading(false)
    }
  }

  if (currentView === 'input') {
    return (
      <div className="app-shell">
        <header className="topbar">
          <h1>SITPO 슬라이드 생성기</h1>
        </header>
        <main className="main-panel input-form">
          <section className="content-card">
            <h2>새 슬라이드 요청</h2>
            <div className="grid two">
              <label className="field">
                <span>대상 (학년)</span>
                <select value={grade} onChange={(event) => handleGradeChange(event.target.value)}>
                  {gradeOptions.map((option) => (
                    <option key={option} value={option}>{option}</option>
                  ))}
                </select>
              </label>
              <label className="field">
                <span>과목</span>
                <select value={subject} onChange={(event) => handleSubjectChange(event.target.value)}>
                  {subjectOptions.map((option) => (
                    <option key={option.subject} value={option.subject}>{option.subject}</option>
                  ))}
                </select>
              </label>
              <label className="field">
                <span>단원</span>
                <select value={unit} onChange={(event) => { setUnit(event.target.value); setTopic(event.target.value) }}>
                  {unitOptions.map((option) => (
                    <option key={option} value={option}>{option}</option>
                  ))}
                </select>
              </label>
              <label className="field">
                <span>주제</span>
                <input value={topic} onChange={(event) => setTopic(event.target.value)} placeholder="예: 생태계의 구성 요소" />
              </label>
              <label className="field">
                <span>슬라이드 수</span>
                <input type="number" min={1} max={30} value={slideCount} onChange={(event) => setSlideCount(parseInt(event.target.value, 10) || 1)} placeholder="예: 9" />
              </label>
              <label className="field">
                <span>스타일</span>
                <input value={style} onChange={(event) => setStyle(event.target.value)} placeholder="예: 프리미엄 과학 탐험 노트형" />
              </label>
            </div>
            <button className="primary wide mt-4" onClick={handleGeneratePlan} disabled={loading}>
              생성 요청 ({loading ? '요청 중' : 'Codex 작업 시작'})
            </button>
            <p className="helper-text">실제 Codex CLI가 슬라이드 설계, 네이티브 이미지 생성, PPTX 조립을 수행합니다. 작업은 몇 분 정도 걸릴 수 있습니다.</p>
            {statusMessage && <p className="helper-text">{statusMessage}</p>}
          </section>
        </main>
      </div>
    )
  }

  return (
    <div className="app-shell">
      <header className="topbar">
        <div>
          <p className="eyebrow">SITPO Slide Generator</p>
          <h1>{project?.title || `${grade} ${subject} - ${topic}`}</h1>
        </div>
        <div className="topbar-actions">
          <button className="ghost" onClick={() => project && downloadText(`${project.id}_plan.md`, exportMarkdown(project))} disabled={!project}>계획서 Markdown</button>
          <button className="ghost" onClick={() => project && downloadText(`${project.id}_project.json`, JSON.stringify(project, null, 2), 'application/json;charset=utf-8')} disabled={!project}>프로젝트 JSON</button>
          <button className="primary" onClick={handleGeneratePptx} disabled={loading || (!project && !serverPptxFile(job))}>PPTX 다운로드</button>
          <button className="primary" onClick={handleProceed} disabled={loading || !job}>[SITPO] 진행</button>
        </div>
      </header>

      <div className="workspace">
        <aside className="stepnav">
          <div className="progress-card">
            <strong>{progressPercent}%</strong>
            <span>작업 상태: {jobLabel(job)}</span>
            <div className="progress-track"><i style={{ width: `${progressPercent}%` }} /></div>
          </div>
          {workflowSteps.map((step, index) => {
            const status = job?.status === 'succeeded' ? '완료' : statusFor(index, currentStepIndex)
            return (
              <div className={`step ${index === currentStepIndex ? 'active' : ''}`} key={step.id}>
                <span className="step-index">{index + 1}</span>
                <span><b>{step.label}</b><small>{step.description}</small></span>
                <em className={`status status-${status.replaceAll(' ', '-')}`}>{status}</em>
              </div>
            )
          })}
        </aside>

        <main className="main-panel">
          <section className="stage-head">
            <div>
              <p className="eyebrow">현재 작업</p>
              <h2>{project ? currentStep.label : jobLabel(job)}</h2>
              <p>{statusMessage || latestLog(job) || 'Codex 작업 상태를 기다리고 있습니다.'}</p>
            </div>
            <div className="stage-actions">
              <button className="secondary" onClick={handleGeneratePptx} disabled={loading || (!project && !serverPptxFile(job))}>PPTX 미리보기</button>
              <button className="primary" onClick={handleProceed} disabled={loading || !job}>[SITPO] 진행</button>
            </div>
          </section>
          <section className="content-card">
            {project ? renderStepContent(project, currentStep, job) : <JobStatusPanel job={job} />}
          </section>
        </main>

        <aside className="inspector">
          <div className="inspector-card accent">
            <p className="eyebrow">연동 상태</p>
            <h3>{statusMessage || jobLabel(job)}</h3>
            <p>서버가 Codex CLI 작업을 실행합니다. 네이티브 이미지 생성이 불가능하면 작업은 명확히 실패로 표시됩니다.</p>
          </div>

          <div className="inspector-card">
            <p className="eyebrow">작업 정보</p>
            <dl className="payload-meta">
              {job && (
                <>
                  <div><dt>Job ID</dt><dd>{job.id}</dd></div>
                  <div><dt>상태</dt><dd>{jobLabel(job)}</dd></div>
                  <div><dt>요청</dt><dd>{job.request.grade} {job.request.subject} · {job.request.topic}</dd></div>
                  <div><dt>슬라이드</dt><dd>{job.request.slideCount}장</dd></div>
                  <div><dt>파일</dt><dd>{job.files.length}개</dd></div>
                </>
              )}
            </dl>
            <button
              className="wide"
              onClick={async () => {
                if (!job) return
                await navigator.clipboard.writeText(JSON.stringify(job, null, 2))
                setStatusMessage('작업 상태 JSON 클립보드 복사 완료')
              }}
              disabled={!job}
            >작업 상태 복사</button>
          </div>

          <div className="inspector-card small-code">
            <pre>{job ? job.logs.slice(-8).join('\n') : '생성 요청 대기 중'}</pre>
          </div>

          {job && <FilesPanel files={job.files} />}
        </aside>
      </div>
    </div>
  )
}

function JobStatusPanel({ job }: { job: SitpoJob | null }) {
  return (
    <div className="job-status">
      <article className="summary-box">
        <span>실제 작업 상태</span>
        <strong>{jobLabel(job)}</strong>
        <p>{latestLog(job) || '생성 요청을 보내면 서버 작업 로그가 여기에 표시됩니다.'}</p>
      </article>
      {job?.error && (
        <article className="summary-box error-box">
          <span>오류</span>
          <strong>작업 실패</strong>
          <p>{job.error}</p>
        </article>
      )}
      <article className="summary-box">
        <span>대기 안내</span>
        <strong>Codex 작업 중</strong>
        <p>서버가 ChatGPT OAuth로 인증된 Codex CLI를 실행하며, 네이티브 이미지 생성까지 검증합니다.</p>
      </article>
    </div>
  )
}

function ResearchPanel({ project }: { project: SitpoProject }) {
  return (
    <div className="grid two">
      <article className="summary-box">
        <span>대상</span>
        <strong>{project.grade} · {project.subject}</strong>
        <p>{project.unit} / {project.topic}</p>
      </article>
      <article className="summary-box">
        <span>스타일</span>
        <strong>{project.style}</strong>
        <p>Codex가 요청값에 맞춰 생성한 실제 프로젝트입니다.</p>
      </article>
      <article className="summary-box">
        <span>슬라이드</span>
        <strong>{project.slides.length}장</strong>
        <p>계획, 활동, 이미지/도식 계획을 포함합니다.</p>
      </article>
      <article className="summary-box">
        <span>현재 상태</span>
        <strong>{project.currentStep}</strong>
        <p>서버 결과 JSON과 PPTX 파일을 확인할 수 있습니다.</p>
      </article>
    </div>
  )
}

function PlanTable({ project }: { project: SitpoProject }) {
  const [selectedSlide, setSelectedSlide] = useState(0)
  const activeSlide = project.slides[selectedSlide] ?? project.slides[0]

  if (!activeSlide) return <JobStatusPanel job={null} />

  return (
    <div className="plan-editor">
      <div className="table-wrap">
        <table>
          <thead>
            <tr><th>장</th><th>제목</th><th>핵심 메시지</th><th>활동</th><th>이미지/도식 계획</th></tr>
          </thead>
          <tbody>
            {project.slides.map((slide, index) => (
              <tr key={slide.slideNo} className={index === selectedSlide ? 'selected' : ''} onClick={() => setSelectedSlide(index)}>
                <td>{slide.slideNo}</td>
                <td><b>{slide.title}</b><small>{slide.learningGoal}</small></td>
                <td>{slide.mainMessage}</td>
                <td>{slide.studentActivity}</td>
                <td><small>이미지: {slide.imagePlan}</small><small>도식: {slide.diagramPlan}</small></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="slide-detail">
        <h3>{activeSlide.title}</h3>
        <p>{activeSlide.mainMessage}</p>
        <p><b>화면 문구:</b> {activeSlide.visibleText.join(' / ')}</p>
        <p><b>학생 활동:</b> {activeSlide.studentActivity}</p>
        <p><b>이미지 계획:</b> {activeSlide.imagePlan}</p>
        <p><b>도식 계획:</b> {activeSlide.diagramPlan}</p>
        <p><b>교사용 메모:</b> {activeSlide.teacherNote}</p>
      </div>
    </div>
  )
}

function AssetBoard({ project }: { project: SitpoProject }) {
  return (
    <div className="cards-list">
      {project.assets.map((asset) => (
        <article className="work-item" key={asset.id}>
          <div><b>{asset.name}</b><span>{asset.kind} · 슬라이드 {asset.slides.join(', ')}</span></div>
          <p>{asset.prompt}</p>
          <em className={`status status-${asset.status.replaceAll(' ', '-')}`}>{asset.status}</em>
        </article>
      ))}
    </div>
  )
}

function DiagramBoard({ project }: { project: SitpoProject }) {
  return (
    <div className="diagram-grid">
      {project.diagrams.map((diagram) => (
        <article className="diagram-card" key={diagram.id}>
          <div className="diagram-top"><b>{diagram.title}</b><span>{diagram.type}</span></div>
          <p>{diagram.nodes.join(' → ')}</p>
          <small>레이아웃: {diagram.layout}</small>
          <small>QA: {diagram.qaRule}</small>
          <em className={`status status-${diagram.status.replaceAll(' ', '-')}`}>{diagram.status}</em>
        </article>
      ))}
    </div>
  )
}

function AssemblyPreview({ project }: { project: SitpoProject }) {
  const [selectedSlide, setSelectedSlide] = useState(0)
  const activeSlide = project.slides[selectedSlide] ?? project.slides[0]
  if (!activeSlide) return <JobStatusPanel job={null} />

  return (
    <div className="assembly">
      <div className="slide-list">
        {project.slides.map((slide, index) => <button key={slide.slideNo} onClick={() => setSelectedSlide(index)}>{slide.slideNo}. {slide.title}</button>)}
      </div>
      <div className="slide-preview">
        <p className="eyebrow">16:9 Preview</p>
        <h3>{activeSlide.title}</h3>
        <p>{activeSlide.mainMessage}</p>
        <ul>{activeSlide.visibleText.map((text) => <li key={text}>{text}</li>)}</ul>
        <div className="visual-plan"><span>{activeSlide.imagePlan}</span><span>{activeSlide.diagramPlan}</span></div>
      </div>
    </div>
  )
}

function QaPanel({ project }: { project: SitpoProject }) {
  return (
    <div className="qa-list">
      {project.qa.map((item) => (
        <label className="qa-item" key={item.id}>
          <input type="checkbox" checked={item.passed} readOnly />
          <span><b>{item.label}</b><small>{item.detail}</small></span>
        </label>
      ))}
    </div>
  )
}

function FilesPanel({ files }: { files: SitpoJobFile[] }) {
  return (
    <div className="inspector-card">
      <p className="eyebrow">서버 파일</p>
      <div className="file-list">
        {files.length === 0 && <small>아직 생성된 파일이 없습니다.</small>}
        {files.map((file) => (
          <a key={file.filename} href={file.url} target="_blank" rel="noreferrer">
            <b>{file.filename}</b>
            <small>{Math.max(1, Math.round(file.sizeBytes / 1024))} KB · {file.mimeType}</small>
          </a>
        ))}
      </div>
    </div>
  )
}

function DownloadPanel({ project, job }: { project: SitpoProject; job: SitpoJob | null }) {
  const pptxFile = serverPptxFile(job)

  return (
    <div className="download-grid">
      <button onClick={() => downloadText(`${project.id}_plan.md`, exportMarkdown(project))}>계획서 Markdown 다운로드</button>
      <button onClick={() => downloadText(`${project.id}_project.json`, JSON.stringify(project, null, 2), 'application/json;charset=utf-8')}>프로젝트 JSON 다운로드</button>
      <button onClick={() => job && downloadText(`${project.id}_job.json`, JSON.stringify(job, null, 2), 'application/json;charset=utf-8')}>작업 상태 JSON 다운로드</button>
      {pptxFile && <a className="download-button primary wide" href={pptxFile.url}>서버 PPTX 다운로드</a>}
      <article>
        <b>서버 산출물</b>
        <p>서버 PPTX가 있으면 우선 사용합니다. 브라우저 PPTX는 서버 파일이 없을 때 보조 미리보기로만 사용합니다.</p>
      </article>
    </div>
  )
}

function renderStepContent(project: SitpoProject, currentStep: WorkflowStep, job: SitpoJob | null) {
  switch (currentStep.id) {
    case 'research':
      return <ResearchPanel project={project} />
    case 'plan':
      return <PlanTable project={project} />
    case 'assets':
      return <AssetBoard project={project} />
    case 'diagrams':
      return <DiagramBoard project={project} />
    case 'assembly':
      return <AssemblyPreview project={project} />
    case 'qa':
      return <QaPanel project={project} />
    case 'download':
      return <DownloadPanel project={project} job={job} />
    default:
      return <div>내용을 로드할 수 없습니다.</div>
  }
}

export default App
