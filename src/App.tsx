import { useMemo, useState } from 'react'
import './App.css'
import { sampleProject, workflowSteps } from './data/sampleProject'
import { createHandoffPayload, downloadText, exportMarkdown } from './lib/exporters'
import type { HandoffPayload, SitpoProject, WorkflowStatus } from './types'

function statusFor(index: number, currentIndex: number): WorkflowStatus {
  if (index < currentIndex) return '완료'
  if (index === currentIndex) return '진행 중'
  return '대기'
}

function App() {
  const [project, setProject] = useState<SitpoProject>(sampleProject)
  const [currentIndex, setCurrentIndex] = useState(0)
  const [selectedSlide, setSelectedSlide] = useState(0)
  const [webhookUrl, setWebhookUrl] = useState(() => localStorage.getItem('sitpo-webhook-url') ?? '')
  const [lastMode, setLastMode] = useState<HandoffPayload['mode']>('continue_step')
  const [toast, setToast] = useState('진행 대기 중')

  const currentStep = workflowSteps[currentIndex]
  const payload = useMemo(
    () => createHandoffPayload(project, currentStep, lastMode, webhookUrl),
    [project, currentStep, lastMode, webhookUrl],
  )

  const progress = Math.round(((currentIndex + 1) / workflowSteps.length) * 100)
  const activeSlide = project.slides[selectedSlide]

  function persistWebhook(value: string) {
    setWebhookUrl(value)
    localStorage.setItem('sitpo-webhook-url', value)
  }

  function markLinkedWork(mode: HandoffPayload['mode']) {
    setLastMode(mode)
    setProject((prev) => ({
      ...prev,
      currentStep: currentStep.id,
      assets: prev.assets.map((asset) =>
        currentStep.id === 'assets' ? { ...asset, status: '진행 중' } : asset,
      ),
      diagrams: prev.diagrams.map((diagram) =>
        currentStep.id === 'diagrams' ? { ...diagram, status: '진행 중' } : diagram,
      ),
    }))
  }

  function handleProgress() {
    markLinkedWork('continue_step')
    if (currentIndex < workflowSteps.length - 1) {
      const nextIndex = currentIndex + 1
      setCurrentIndex(nextIndex)
      setToast(`진행 연결됨: ${workflowSteps[nextIndex].label} 단계로 이동`)
    } else {
      setToast('모든 단계가 완료 상태입니다. 다운로드/공유를 진행하세요.')
    }
  }

  function approveAndBuild() {
    markLinkedWork('approve_and_build')
    const assemblyIndex = workflowSteps.findIndex((step) => step.id === 'assembly')
    setCurrentIndex(Math.max(assemblyIndex, currentIndex))
    setToast('승인 후 제작 모드로 작업 페이로드가 갱신되었습니다.')
  }

  async function copyPayload() {
    const text = JSON.stringify(payload, null, 2)
    try {
      await navigator.clipboard.writeText(text)
      setToast('작업 페이로드를 클립보드에 복사했습니다.')
    } catch {
      downloadText('sitpo_handoff_payload.json', text, 'application/json;charset=utf-8')
      setToast('클립보드가 막혀 JSON 파일로 다운로드했습니다.')
    }
  }

  function toggleQa(id: string) {
    setProject((prev) => ({
      ...prev,
      qa: prev.qa.map((item) => (item.id === id ? { ...item, passed: !item.passed } : item)),
    }))
  }

  function renderStepContent() {
    if (currentStep.id === 'research') {
      return <ResearchPanel project={project} />
    }
    if (currentStep.id === 'plan') {
      return <PlanTable project={project} selectedSlide={selectedSlide} setSelectedSlide={setSelectedSlide} />
    }
    if (currentStep.id === 'assets') {
      return <AssetBoard project={project} />
    }
    if (currentStep.id === 'diagrams') {
      return <DiagramBoard project={project} />
    }
    if (currentStep.id === 'assembly') {
      return <AssemblyPreview project={project} activeSlide={activeSlide} setSelectedSlide={setSelectedSlide} />
    }
    if (currentStep.id === 'qa') {
      return <QaPanel project={project} toggleQa={toggleQa} />
    }
    return <DownloadPanel project={project} payload={payload} />
  }

  return (
    <div className="app-shell">
      <header className="topbar">
        <div>
          <p className="eyebrow">SITPO Slide Workflow MVP</p>
          <h1>{project.title}</h1>
        </div>
        <div className="topbar-actions">
          <button className="ghost" onClick={() => downloadText('sitpo_slide_plan.md', exportMarkdown(project))}>Markdown</button>
          <button className="ghost" onClick={() => downloadText('sitpo_project.json', JSON.stringify(project, null, 2), 'application/json;charset=utf-8')}>JSON</button>
          <button className="primary" onClick={handleProgress}>[SITPO] 진행</button>
        </div>
      </header>

      <div className="workspace">
        <aside className="stepnav">
          <div className="progress-card">
            <strong>{progress}%</strong>
            <span>현재 단계: {currentStep.label}</span>
            <div className="progress-track"><i style={{ width: `${progress}%` }} /></div>
          </div>
          {workflowSteps.map((step, index) => {
            const status = statusFor(index, currentIndex)
            return (
              <button
                className={`step ${index === currentIndex ? 'active' : ''}`}
                key={step.id}
                onClick={() => setCurrentIndex(index)}
              >
                <span className="step-index">{index + 1}</span>
                <span><b>{step.label}</b><small>{step.description}</small></span>
                <em className={`status status-${status.replaceAll(' ', '-')}`}>{status}</em>
              </button>
            )
          })}
        </aside>

        <main className="main-panel">
          <section className="stage-head">
            <div>
              <p className="eyebrow">현재 작업</p>
              <h2>{currentStep.label}</h2>
              <p>{currentStep.description}</p>
            </div>
            <div className="stage-actions">
              <button className="secondary" onClick={approveAndBuild}>승인 후 제작</button>
              <button className="primary" onClick={handleProgress}>진행 연결</button>
            </div>
          </section>
          <section className="content-card">{renderStepContent()}</section>
        </main>

        <aside className="inspector">
          <div className="inspector-card accent">
            <p className="eyebrow">연동 상태</p>
            <h3>{toast}</h3>
            <p>버튼을 누르면 현재 단계, 다음 액션, 슬라이드/에셋/도식 데이터가 작업 페이로드로 즉시 갱신됩니다.</p>
          </div>

          <label className="field">
            <span>Hermes/Webhook URL</span>
            <input
              placeholder="선택 입력: https://.../webhooks/sitpo"
              value={webhookUrl}
              onChange={(event) => persistWebhook(event.target.value)}
            />
          </label>

          <div className="inspector-card">
            <p className="eyebrow">현재 페이로드</p>
            <dl className="payload-meta">
              <div><dt>명령</dt><dd>{payload.command}</dd></div>
              <div><dt>모드</dt><dd>{payload.mode}</dd></div>
              <div><dt>다음 액션</dt><dd>{payload.nextAction}</dd></div>
              <div><dt>슬라이드</dt><dd>{payload.slideCount}장</dd></div>
              <div><dt>QA</dt><dd>{payload.qaPassed}/{payload.qaTotal}</dd></div>
            </dl>
            <button className="wide" onClick={copyPayload}>작업 페이로드 복사</button>
          </div>

          <div className="inspector-card small-code">
            <pre>{JSON.stringify({ command: payload.command, step: payload.currentStep, nextAction: payload.nextAction }, null, 2)}</pre>
          </div>
        </aside>
      </div>
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
        <p>딥그린·크림·라임 포인트, 교사용 실용 도구형 UI</p>
      </article>
      <article className="summary-box">
        <span>핵심 개념</span>
        <strong>생산자 · 소비자 · 분해자 · 비생물 요소</strong>
        <p>관찰 → 분류 → 관계 → 균형 변화 순서로 수업 흐름을 구성합니다.</p>
      </article>
      <article className="summary-box">
        <span>바로 연동</span>
        <strong>[SITPO] 진행</strong>
        <p>진행 버튼이 작업 상태와 Hermes/Codex 전달용 JSON을 함께 갱신합니다.</p>
      </article>
    </div>
  )
}

function PlanTable({ project, selectedSlide, setSelectedSlide }: { project: SitpoProject; selectedSlide: number; setSelectedSlide: (index: number) => void }) {
  return (
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
        </article>
      ))}
    </div>
  )
}

function AssemblyPreview({ project, activeSlide, setSelectedSlide }: { project: SitpoProject; activeSlide: SitpoProject['slides'][number]; setSelectedSlide: (index: number) => void }) {
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

function QaPanel({ project, toggleQa }: { project: SitpoProject; toggleQa: (id: string) => void }) {
  return (
    <div className="qa-list">
      {project.qa.map((item) => (
        <label className="qa-item" key={item.id}>
          <input type="checkbox" checked={item.passed} onChange={() => toggleQa(item.id)} />
          <span><b>{item.label}</b><small>{item.detail}</small></span>
        </label>
      ))}
    </div>
  )
}

function DownloadPanel({ project, payload }: { project: SitpoProject; payload: HandoffPayload }) {
  return (
    <div className="download-grid">
      <button onClick={() => downloadText('sitpo_slide_plan.md', exportMarkdown(project))}>계획서 Markdown 다운로드</button>
      <button onClick={() => downloadText('sitpo_project.json', JSON.stringify(project, null, 2), 'application/json;charset=utf-8')}>프로젝트 JSON 다운로드</button>
      <button onClick={() => downloadText('sitpo_handoff_payload.json', JSON.stringify(payload, null, 2), 'application/json;charset=utf-8')}>작업 페이로드 다운로드</button>
      <article>
        <b>다음 서버 연동 지점</b>
        <p>Webhook URL이 들어오면 이 페이로드를 Hermes/Codex 작업 큐로 POST하는 API만 추가하면 됩니다.</p>
      </article>
    </div>
  )
}

export default App
