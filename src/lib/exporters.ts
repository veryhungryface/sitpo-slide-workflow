import type { HandoffPayload, SitpoProject, WorkflowStep } from '../types'

export function createHandoffPayload(
  project: SitpoProject,
  currentStep: WorkflowStep,
  mode: HandoffPayload['mode'],
  webhookUrl?: string,
): HandoffPayload {
  return {
    source: 'sitpo-slide-workflow-mvp',
    command: '[SITPO] 진행',
    mode,
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
}

export function exportMarkdown(project: SitpoProject): string {
  const slideRows = project.slides
    .map((slide) => `### ${slide.slideNo}. ${slide.title}\n- 학습목표: ${slide.learningGoal}\n- 핵심 메시지: ${slide.mainMessage}\n- 화면 문구: ${slide.visibleText.join(' / ')}\n- 학생 활동: ${slide.studentActivity}\n- 이미지 계획: ${slide.imagePlan}\n- 도식 계획: ${slide.diagramPlan}\n- 교사용 메모: ${slide.teacherNote}`)
    .join('\n\n')

  const assetRows = project.assets
    .map((asset) => `- ${asset.name} (${asset.kind}) — 슬라이드 ${asset.slides.join(', ')} — ${asset.status}`)
    .join('\n')

  const diagramRows = project.diagrams
    .map((diagram) => `- ${diagram.title} [${diagram.type}] — ${diagram.layout} — QA: ${diagram.qaRule}`)
    .join('\n')

  const qaRows = project.qa
    .map((item) => `- [${item.passed ? 'x' : ' '}] ${item.label}: ${item.detail}`)
    .join('\n')

  return `# ${project.title}\n\n## 프로젝트\n- 대상: ${project.grade}\n- 과목/단원: ${project.subject} / ${project.unit}\n- 주제: ${project.topic}\n- 스타일: ${project.style}\n\n## 슬라이드 계획\n\n${slideRows}\n\n## 이미지/에셋 생성 계획\n${assetRows}\n\n## 도식 생성 계획\n${diagramRows}\n\n## QA 체크리스트\n${qaRows}\n`
}

export function downloadText(filename: string, content: string | Blob, mime = 'text/plain;charset=utf-8') {
  const blob = typeof content === 'string' ? new Blob([content], { type: mime }) : content
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  document.body.appendChild(link)
  link.click()
  link.remove()
  URL.revokeObjectURL(url)
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
