import { useState, useMemo } from 'react'
import './App.css'
import { getSubjectsForGrade, getUnitsForGradeSubject, gradeOptions } from './data/elementaryCurriculum'
import { sampleProject, workflowSteps } from './data/sampleProject'
import { downloadText, exportMarkdown } from './lib/exporters'
import { generatePptx } from './lib/pptxGenerator'
import type { HandoffPayload, SitpoProject, WorkflowStatus, SlidePlan, AssetPlan, DiagramSpec, WorkflowStep } from './types'

// --- Helper Functions --- //
function statusFor(index: number, currentIndex: number): WorkflowStatus {
  if (index < currentIndex) return '완료'
  if (index === currentIndex) return '진행 중'
  return '대기'
}

function nextActionForStep(stepId: string) {
  switch (stepId) {
    case 'research':
      return '교육과정·핵심 개념 조사 후 슬라이드 계획서 작성'
    case 'plan':
      return '계획서 승인 후 이미지/도식 작업 단위 생성'
    case 'assets':
      return 'Codex 네이티브 이미지 생성 및 투명 PNG 후처리'
    case 'diagrams':
      return 'DiagramSpec를 SVG/PNG로 안정 렌더링'
    case 'assembly':
      return 'PptxGenJS로 슬라이드 조립 및 미리보기 생성'
    case 'qa':
      return 'PDF/PNG 렌더링 QA 후 수정 루프 실행'
    default:
      return '산출물 다운로드 및 공유 링크 생성'
  }
}

const activityTemplates: Record<string, string[]> = {
  국어: ['핵심 문장 찾기', '생각을 근거와 함께 말하기', '짧은 글로 정리하기'],
  수학: ['개념 카드 분류하기', '대표 문제 함께 해결하기', '내 풀이 과정을 설명하기'],
  사회: ['자료를 보고 사실 찾기', '사례를 비교해 보기', '내 생활과 연결해 말하기'],
  과학: ['관찰 결과 기록하기', '원인과 결과 연결하기', '개념 모형으로 설명하기'],
  영어: ['핵심 표현 따라 말하기', '짝과 묻고 답하기', '짧은 문장 만들기'],
  음악: ['리듬이나 가락 따라 표현하기', '느낌을 말로 나누기', '간단한 창작 활동하기'],
  미술: ['이미지 관찰하기', '표현 방법 선택하기', '작품 의도 설명하기'],
  체육: ['동작 순서 익히기', '안전 규칙 확인하기', '팀 활동 후 피드백하기'],
  실과: ['생활 사례 찾기', '절차를 순서대로 정리하기', '간단한 실습 계획 세우기'],
  도덕: ['상황 판단하기', '내 선택의 이유 말하기', '실천 다짐 쓰기'],
}

function buildSlidesFromRequest(params: { slideCount: number; subject: string; unit: string; topic: string }): SlidePlan[] {
  const stages = ['도입', '핵심 개념', '예시 탐구', '방법 익히기', '함께 연습', '자료 분석', '오개념 점검', '적용 활동', '정리와 성찰']
  const activities = activityTemplates[params.subject] ?? ['핵심 내용 확인하기', '예시와 비예시 비교하기', '한 문장으로 정리하기']

  return Array.from({ length: params.slideCount }, (_, index) => {
    const stage = stages[index % stages.length]
    const slideNo = index + 1

    return {
      slideNo,
      title: `${params.topic} - ${slideNo}차시`,
      learningGoal: `${params.subject} ${params.unit} 단원에서 ${params.topic}의 핵심을 이해합니다.`,
      mainMessage: `${stage} 단계에서는 ${params.topic}을 학생 활동과 시각 자료로 연결합니다.`,
      visibleText: [params.unit, params.topic, stage],
      studentActivity: activities[index % activities.length],
      imagePlan: `${params.subject} ${params.topic} 이해를 돕는 16:9 수업 장면 또는 핵심 예시 이미지`,
      diagramPlan: `${params.topic} 핵심 관계·절차·비교 구조를 한눈에 보는 도식`,
      teacherNote: `선택한 과목/단원(${params.subject} · ${params.unit}) 기준으로 실제 수업 흐름에 맞춰 구체화합니다.`,
    }
  })
}

// --- Mocking AI/Codex Generation --- //
const simulateAiGeneration = (request: { prompt: string; type: 'slides' | 'assets' | 'diagrams' }): Promise<Partial<SitpoProject>> => {
  return new Promise((resolve) => {
    setTimeout(() => {
      if (request.type === 'slides') {
        const generatedSlides: SlidePlan[] = sampleProject.slides.map((slide) => ({ ...slide, imagePlan: '[생성 이미지 적용 예정]', diagramPlan: '[생성 도식 적용 예정]' }))
        resolve({ slides: generatedSlides, currentStep: 'assembly' })
      } else if (request.type === 'assets') {
        const generatedAssets: AssetPlan[] = sampleProject.assets.map((asset) => ({ ...asset, status: '완료' }))
        resolve({ assets: generatedAssets, currentStep: 'assets' })
      } else if (request.type === 'diagrams') {
        const generatedDiagrams: DiagramSpec[] = sampleProject.diagrams.map((diagram) => ({ ...diagram, status: '완료' }))
        resolve({ diagrams: generatedDiagrams, currentStep: 'diagrams' })
      }
      resolve({})
    }, 2000) // Simulate network delay
  })
}

// --- Main App Component --- //
function App() {
  const [project, setProject] = useState<SitpoProject | null>(null)
  const [currentView, setCurrentView] = useState<'input' | 'plan' | 'preview'>('input')
  const [grade, setGrade] = useState('초5')
  const [subject, setSubject] = useState('과학')
  const [unit, setUnit] = useState('생물과 환경')
  const [topic, setTopic] = useState('생태계의 구성 요소')
  const [slideCount, setSlideCount] = useState(9)
  const [style, setStyle] = useState('프리미엄 과학 탐험 노트형')
  const [statusMessage, setStatusMessage] = useState('')
  const [webhookUrl, setWebhookUrl] = useState(() => localStorage.getItem('sitpo-webhook-url') ?? '')
  const [loading, setLoading] = useState(false)

  const subjectOptions = useMemo(() => getSubjectsForGrade(grade), [grade])
  const unitOptions = useMemo(() => getUnitsForGradeSubject(grade, subject), [grade, subject])
  const currentStep = project ? workflowSteps.find(step => step.id === project.currentStep) || workflowSteps[0] : workflowSteps[0]

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

  const handleUnitChange = (nextUnit: string) => {
    setUnit(nextUnit)
    setTopic(nextUnit)
  }

  const handoffPayload: HandoffPayload | null = useMemo(() => {
    if (!project) return null
    return {
      source: 'sitpo-slide-workflow-mvp',
      command: '[SITPO] 진행',
      mode: 'continue_step',
      webhookUrl: webhookUrl?.trim() || undefined,
      currentStep: currentStep.label,
      nextAction: nextActionForStep(currentStep.id),
      project: {
        id: project.id,
        title: project.title,
        grade: project.grade,
        subject: project.subject,
        unit: project.unit,
        topic: project.topic,
        style: project.style,
      },
      slideCount: project.slides.length,
      assetCount: project.assets.length,
      diagramCount: project.diagrams.length,
      qaPassed: project.qa.filter((item) => item.passed).length,
      qaTotal: project.qa.length,
      requestedOutputs: ['slide_plan.md', 'sitpo_project.json', 'handoff_payload.json', 'PPTX', 'PDF'],
      slides: project.slides,
      assets: project.assets,
      diagrams: project.diagrams,
    }
  }, [project, currentStep, webhookUrl])

  const handleGeneratePlan = async () => {
    setLoading(true)
    setStatusMessage('슬라이드 계획서를 생성 중입니다...')
    // Simulate plan generation based on inputs
    const newProject: SitpoProject = {
      ...sampleProject,
      id: `sitpo-${Date.now()}`,
      grade, subject, unit, topic, style,
      slides: buildSlidesFromRequest({ slideCount, subject, unit, topic }),
      currentStep: 'plan',
      title: `${grade} ${subject} - ${topic}`,
    }
    setProject(newProject)
    setCurrentView('plan')
    setStatusMessage('계획서 생성이 완료되었습니다.')
    setLoading(false)
  }

  const postHandoffPayload = async (payload: HandoffPayload) => {
    const targetUrl = webhookUrl.trim()
    if (!targetUrl) return false

    const response = await fetch(targetUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    })

    if (!response.ok) {
      throw new Error(`Webhook failed: ${response.status}`)
    }

    return true
  }

  const handleProceed = async () => {
    if (!project) return

    const nextStepIndex = workflowSteps.findIndex(step => step.id === project.currentStep) + 1
    if (nextStepIndex < workflowSteps.length) {
      const nextStep = workflowSteps[nextStepIndex]
      setLoading(true)
      setStatusMessage(`${nextStep.label} 단계 작업을 진행 중입니다...`)

      try {
        const sentToWebhook = handoffPayload ? await postHandoffPayload(handoffPayload) : false
        let updatedProject = { ...project }

        if (nextStep.id === 'assets') {
          const res = await simulateAiGeneration({ prompt: `Generate assets for ${project.topic}`, type: 'assets' })
          updatedProject = { ...updatedProject, ...res, currentStep: 'assets' }
        } else if (nextStep.id === 'diagrams') {
          const res = await simulateAiGeneration({ prompt: `Generate diagrams for ${project.topic}`, type: 'diagrams' })
          updatedProject = { ...updatedProject, ...res, currentStep: 'diagrams' }
        } else if (nextStep.id === 'assembly') {
          const res = await simulateAiGeneration({ prompt: `Assemble slides for ${project.topic}`, type: 'slides' })
          updatedProject = { ...updatedProject, ...res, currentStep: 'assembly' }
        } else if (nextStep.id === 'qa') {
          updatedProject = { ...updatedProject, qa: updatedProject.qa.map(item => ({ ...item, passed: true })), currentStep: 'qa' }
        } else {
          updatedProject = { ...updatedProject, currentStep: nextStep.id }
        }

        setProject(updatedProject)
        setStatusMessage(`${sentToWebhook ? '웹훅 전송 완료 · ' : ''}${nextStep.label} 단계 완료. 다음: ${workflowSteps[nextStepIndex + 1]?.label || '다운로드'}`)
        setCurrentView('preview')
      } catch (error) {
        console.error('SITPO handoff failed:', error)
        setStatusMessage('웹훅 전송에 실패했습니다. URL/CORS 설정을 확인해 주세요.')
      } finally {
        setLoading(false)
      }
    } else {
      setStatusMessage('모든 워크플로우 단계가 완료되었습니다.')
    }
  }

  const handleGeneratePptx = async () => {
    if (!project) return
    setStatusMessage('PPTX 파일을 생성 중입니다...')
    setLoading(true)
    try {
      const pptxBlob = await generatePptx(project)
      downloadText(`${project.id}.pptx`, pptxBlob, 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
      setStatusMessage('PPTX 파일 다운로드를 시작했습니다.')
    } catch (error) {
      console.error('Error generating PPTX:', error)
      setStatusMessage('PPTX 생성 중 오류가 발생했습니다.')
    } finally {
      setLoading(false)
    }
  }

  function persistWebhook(value: string) {
    setWebhookUrl(value)
    localStorage.setItem('sitpo-webhook-url', value)
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
                <select value={grade} onChange={(e) => handleGradeChange(e.target.value)}>
                  {gradeOptions.map((option) => (
                    <option key={option} value={option}>{option}</option>
                  ))}
                </select>
              </label>
              <label className="field">
                <span>과목</span>
                <select value={subject} onChange={(e) => handleSubjectChange(e.target.value)}>
                  {subjectOptions.map((option) => (
                    <option key={option.subject} value={option.subject}>{option.subject}</option>
                  ))}
                </select>
              </label>
              <label className="field">
                <span>단원</span>
                <select value={unit} onChange={(e) => handleUnitChange(e.target.value)}>
                  {unitOptions.map((option) => (
                    <option key={option} value={option}>{option}</option>
                  ))}
                </select>
              </label>
              <label className="field">
                <span>주제</span>
                <input value={topic} onChange={(e) => setTopic(e.target.value)} placeholder="예: 생태계의 구성 요소" />
              </label>
              <label className="field">
                <span>슬라이드 수</span>
                <input type="number" value={slideCount} onChange={(e) => setSlideCount(parseInt(e.target.value))} placeholder="예: 9" />
              </label>
              <label className="field">
                <span>스타일</span>
                <input value={style} onChange={(e) => setStyle(e.target.value)} placeholder="예: 프리미엄 과학 탐험 노트형" />
              </label>
            </div>
            <button className="primary wide mt-4" onClick={handleGeneratePlan} disabled={loading}>생성 요청 ({loading ? statusMessage : '클릭'})</button>
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
          <h1>{project?.title || '슬라이드 프로젝트'}</h1>
        </div>
        <div className="topbar-actions">
          <button className="ghost" onClick={() => project && downloadText(`${project.id}_plan.md`, exportMarkdown(project))}>계획서 Markdown</button>
          <button className="ghost" onClick={() => project && downloadText(`${project.id}_project.json`, JSON.stringify(project, null, 2), 'application/json;charset=utf-8')}>프로젝트 JSON</button>
          <button className="primary" onClick={handleGeneratePptx} disabled={loading}>PPTX 다운로드</button>
          <button className="primary" onClick={handleProceed} disabled={loading}>[SITPO] 진행 ({workflowSteps.findIndex(step => step.id === project?.currentStep) + 1}/{workflowSteps.length})</button>
        </div>
      </header>

      <div className="workspace">
        <aside className="stepnav">
          <div className="progress-card">
            <strong>{Math.round(((workflowSteps.findIndex(step => step.id === project?.currentStep) + 1) / workflowSteps.length) * 100)}%</strong>
            <span>현재 단계: {currentStep.label}</span>
            <div className="progress-track"><i style={{ width: `${Math.round(((workflowSteps.findIndex(step => step.id === project?.currentStep) + 1) / workflowSteps.length) * 100)}%` }} /></div>
          </div>
          {workflowSteps.map((step, index) => {
            const currentStepIndex = workflowSteps.findIndex(s => s.id === project?.currentStep)
            const status = statusFor(index, currentStepIndex)
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
              <h2>{currentStep.label}</h2>
              <p>{statusMessage || currentStep.description}</p>
            </div>
            <div className="stage-actions">
              <button className="secondary" onClick={handleGeneratePptx} disabled={loading}>PPTX 미리보기</button>
              <button className="primary" onClick={handleProceed} disabled={loading}>[SITPO] 진행</button>
            </div>
          </section>
          <section className="content-card">{project && renderStepContent(project, currentStep, currentStep.id === 'qa' ? (id) => setProject(prev => prev ? ({ ...prev, qa: prev.qa.map(item => (item.id === id ? { ...item, passed: !item.passed } : item)) }) : null) : undefined)}</section>
        </main>

        <aside className="inspector">
          <div className="inspector-card accent">
            <p className="eyebrow">연동 상태</p>
            <h3>{statusMessage}</h3>
            <p>SITPO 진행 버튼을 누르면 다음 단계로 이동하며, 내부 페이로드가 갱신됩니다. 웹훅 URL로 Hermes/Codex에 연동할 수 있습니다.</p>
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
            <p className="eyebrow">현재 페이로드 미리보기</p>
            <dl className="payload-meta">
              {handoffPayload && (
                <>
                  <div><dt>명령</dt><dd>{handoffPayload.command}</dd></div>
                  <div><dt>모드</dt><dd>{handoffPayload.mode}</dd></div>
                  <div><dt>다음 액션</dt><dd>{handoffPayload.nextAction}</dd></div>
                  <div><dt>슬라이드</dt><dd>{handoffPayload.slideCount}장</dd></div>
                  <div><dt>QA</dt><dd>{handoffPayload.qaPassed}/{handoffPayload.qaTotal}</dd></div>
                </>
              )}
            </dl>
            <button
              className="wide"
              onClick={async () => {
                if (!handoffPayload) return
                await navigator.clipboard.writeText(JSON.stringify(handoffPayload, null, 2))
                setStatusMessage('페이로드 클립보드 복사 완료')
              }}
              disabled={!handoffPayload}
            >작업 페이로드 복사</button>
          </div>

          <div className="inspector-card small-code">
            <pre>{handoffPayload ? JSON.stringify({ command: handoffPayload.command, step: handoffPayload.currentStep, nextAction: handoffPayload.nextAction }, null, 2) : '생성 요청 대기 중'}</pre>
          </div>
        </aside>
      </div>
    </div>
  )
}

// --- Helper Components for each step --- //
// NOTE: In a real app, these would be separate files.

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
        <span>현재 상태</span>
        <strong>{project.currentStep}</strong>
        <p>요청 정보를 바탕으로 계획서를 자동으로 생성했습니다. [SITPO] 진행 버튼으로 다음 단계로 이동합니다.</p>
      </article>
    </div>
  )
}

function PlanTable({ project }: { project: SitpoProject }) {
  const [selectedSlide, setSelectedSlide] = useState(0)
  const activeSlide = project.slides[selectedSlide]

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
  const activeSlide = project.slides[selectedSlide]
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

function DownloadPanel({ project, payload }: { project: SitpoProject; payload: HandoffPayload | null }) {
  return (
    <div className="download-grid">
      <button onClick={() => downloadText(`${project.id}_plan.md`, exportMarkdown(project))}>계획서 Markdown 다운로드</button>
      <button onClick={() => downloadText(`${project.id}_project.json`, JSON.stringify(project, null, 2), 'application/json;charset=utf-8')}>프로젝트 JSON 다운로드</button>
      <button onClick={() => payload && downloadText(`${project.id}_handoff_payload.json`, JSON.stringify(payload, null, 2), 'application/json;charset=utf-8')}>작업 페이로드 다운로드</button>
      <button className="primary wide" onClick={async () => { const pptxBlob = await generatePptx(project); downloadText(`${project.id}.pptx`, pptxBlob, 'application/vnd.openxmlformats-officedocument.presentationml.presentation') }}>PPTX 다운로드</button>
      <article>
        <b>다음 서버 연동 지점</b>
        <p>Webhook URL이 들어오면 이 페이로드를 Hermes/Codex 작업 큐로 POST하는 API만 추가하면 됩니다.</p>
      </article>
    </div>
  )
}

function renderStepContent(project: SitpoProject, currentStep: WorkflowStep, toggleQa?: (id: string) => void) {
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
      return <QaPanel project={project} toggleQa={toggleQa!} />
    case 'download':
      return <DownloadPanel project={project} payload={null} /> // Payload handled by App.tsx
    default:
      return <div>내용을 로드할 수 없습니다.</div>
  }
}

export default App
