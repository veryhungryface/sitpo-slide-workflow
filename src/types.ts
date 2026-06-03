export type WorkflowStatus = '대기' | '진행 중' | '완료' | '수정 필요'
export type SitpoJobStatus = 'queued' | 'running' | 'succeeded' | 'failed'

export type WorkflowStep = {
  id: string
  label: string
  short: string
  description: string
}

export type SlidePlan = {
  slideNo: number
  title: string
  learningGoal: string
  mainMessage: string
  visibleText: string[]
  studentActivity: string
  imagePlan: string
  diagramPlan: string
  teacherNote: string
}

export type AssetPlan = {
  id: string
  name: string
  slides: number[]
  kind: '배경' | '투명 PNG' | '아이콘 세트' | '캐릭터/오브젝트'
  prompt: string
  status: WorkflowStatus
}

export type DiagramSpec = {
  id: string
  title: string
  type: 'process_flow' | 'comparison_table' | 'cycle' | 'relationship_map' | 'classification'
  slides: number[]
  nodes: string[]
  layout: string
  qaRule: string
  status: WorkflowStatus
}

export type QACheckItem = {
  id: string
  label: string
  detail: string
  passed: boolean
}

export type SitpoProject = {
  id: string
  title: string
  grade: string
  subject: string
  unit: string
  topic: string
  style: string
  currentStep: string
  createdAt: string
  slides: SlidePlan[]
  assets: AssetPlan[]
  diagrams: DiagramSpec[]
  qa: QACheckItem[]
}

export type SitpoJobRequest = {
  grade: string
  subject: string
  unit: string
  topic: string
  style: string
  slideCount: number
}

export type SitpoJobFile = {
  filename: string
  url: string
  sizeBytes: number
  mimeType: string
}

export type SitpoJob = {
  id: string
  status: SitpoJobStatus
  request: SitpoJobRequest
  createdAt: string
  updatedAt: string
  logs: string[]
  files: SitpoJobFile[]
  result?: SitpoProject
  error?: string
}

export type HandoffPayload = {
  source: 'sitpo-slide-workflow-mvp'
  command: '[SITPO] 진행'
  mode: 'plan_only' | 'approve_and_build' | 'continue_step'
  webhookUrl?: string
  currentStep: string
  nextAction: string
  project: Pick<SitpoProject, 'id' | 'title' | 'grade' | 'subject' | 'unit' | 'topic' | 'style'>
  slideCount: number
  assetCount: number
  diagramCount: number
  qaPassed: number
  qaTotal: number
  requestedOutputs: string[]
  slides: SlidePlan[]
  assets: AssetPlan[]
  diagrams: DiagramSpec[]
}
