import fs from 'node:fs'
import path from 'node:path'
import PptxGenJS from 'pptxgenjs'

const [, , inputArg, outputArg] = process.argv

if (!inputArg || !outputArg) {
  console.error('Usage: node server/build-pptx.mjs <sitpo_project.json> <output.pptx>')
  process.exit(2)
}

const inputPath = path.resolve(inputArg)
const outputPath = path.resolve(outputArg)
const jobDir = path.dirname(inputPath)
const project = JSON.parse(fs.readFileSync(inputPath, 'utf8'))

const W = 13.333
const H = 7.5
const palette = {
  dark: '12392D',
  dark2: '16352C',
  green: '2F6B4F',
  moss: '8FB56A',
  sage: 'DCEAC8',
  cream: 'F8F2E4',
  paper: 'FFFDF4',
  sand: 'EEDFC0',
  ink: '17352C',
  muted: '537466',
  line: 'CAD9BE',
  accent: 'D98D39',
  coral: 'D65C4F',
  blue: '4C7A91',
}

function firstExistingImage(...candidates) {
  for (const candidate of candidates.filter(Boolean)) {
    const filename = path.basename(String(candidate))
    const fullPath = path.join(jobDir, filename)
    if (fs.existsSync(fullPath) && /\.(png|jpe?g|webp)$/i.test(fullPath)) {
      return fullPath
    }
  }
  return null
}

function assetsForSlide(slideData) {
  const ids = new Set(Array.isArray(slideData.assetIds) ? slideData.assetIds : [])
  return (project.assets || [])
    .filter((asset) => {
      const usedBySlide = Array.isArray(asset.slides) && asset.slides.includes(slideData.slideNo)
      return ids.has(asset.id) || usedBySlide
    })
    .map((asset) => ({
      ...asset,
      imagePath: firstExistingImage(asset.fileName, asset.filename, asset.path, asset.generatedImage),
    }))
    .filter((asset) => asset.imagePath)
}

function sheetImagePath() {
  return firstExistingImage(project.assetSheet?.fileName)
}

function allAssets() {
  return (project.assets || [])
    .map((asset) => ({ ...asset, imagePath: firstExistingImage(asset.fileName, asset.filename, asset.path, asset.generatedImage) }))
    .filter((asset) => asset.imagePath)
}

function safeText(value, fallback = '') {
  return typeof value === 'string' && value.trim() ? value.trim() : fallback
}

function visibleText(slideData) {
  return Array.isArray(slideData.visibleText) ? slideData.visibleText.filter(Boolean) : []
}

function splitItems(items, count) {
  const groups = Array.from({ length: count }, () => [])
  items.forEach((item, index) => groups[index % count].push(item))
  return groups
}

function addPageNumber(slide, slideNo) {
  slide.addText(String(slideNo).padStart(2, '0'), {
    x: 12.25,
    y: 6.88,
    w: 0.65,
    h: 0.18,
    fontSize: 8,
    bold: true,
    color: '6C7E74',
    align: 'right',
    margin: 0,
  })
}

function addTopBand(slide, slideData, options = {}) {
  const dark = options.dark || false
  const bandColor = dark ? palette.dark : palette.cream
  const textColor = dark ? 'FFFFFF' : palette.ink
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: W,
    h: 0.28,
    fill: { color: dark ? palette.accent : palette.dark },
    line: { color: dark ? palette.accent : palette.dark },
  })
  if (dark) {
    slide.background = { color: palette.dark }
  } else {
    slide.background = { color: bandColor }
  }
  slide.addText(`${project.grade} ${project.subject} · ${project.unit}`, {
    x: 0.55,
    y: 0.48,
    w: 7.8,
    h: 0.23,
    fontSize: 8.5,
    bold: true,
    color: dark ? 'CFE3C4' : palette.muted,
    margin: 0,
  })
  slide.addText(safeText(slideData.title, project.title), {
    x: 0.55,
    y: 0.78,
    w: options.titleW || 8.9,
    h: 0.62,
    fontSize: options.titleSize || 26,
    bold: true,
    color: textColor,
    fit: 'shrink',
    margin: 0.02,
  })
}

function addMainMessage(slide, slideData, box) {
  slide.addText(safeText(slideData.mainMessage), {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    fontSize: box.fontSize || 15,
    bold: box.bold || false,
    color: box.color || palette.dark2,
    fit: 'shrink',
    margin: 0.03,
    breakLine: false,
  })
}

function addActivityCard(slide, slideData, box, options = {}) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    rectRadius: 0.08,
    fill: { color: options.fill || palette.sage },
    line: { color: options.line || palette.line },
  })
  slide.addText('활동', {
    x: box.x + 0.2,
    y: box.y + 0.14,
    w: 0.7,
    h: 0.2,
    fontSize: 9,
    bold: true,
    color: options.accent || palette.green,
    margin: 0,
  })
  slide.addText(safeText(slideData.studentActivity, '생각을 말해 보기'), {
    x: box.x + 0.2,
    y: box.y + 0.42,
    w: box.w - 0.4,
    h: Math.max(0.25, box.h - 0.52),
    fontSize: options.fontSize || 12,
    color: options.textColor || palette.ink,
    fit: 'shrink',
    margin: 0.02,
  })
}

function addAsset(slide, asset, box, options = {}) {
  if (!asset?.imagePath) return false
  slide.addImage({ path: asset.imagePath, x: box.x, y: box.y, w: box.w, h: box.h, sizingContain: true })
  if (options.label !== false) {
    slide.addText(safeText(asset.name || asset.id), {
      x: box.x,
      y: box.y + box.h + 0.02,
      w: box.w,
      h: 0.18,
      fontSize: options.labelSize || 7,
      color: options.labelColor || palette.muted,
      align: 'center',
      margin: 0,
      fit: 'shrink',
    })
  }
  return true
}

function pickAssets(slideData, fallbackCount = 4) {
  const slideAssets = assetsForSlide(slideData)
  if (slideAssets.length) return slideAssets
  const assets = allAssets()
  if (assets.length) return assets.slice(0, fallbackCount)
  const sheet = sheetImagePath()
  return sheet ? [{ id: 'asset-sheet', name: '에셋 시트', imagePath: sheet }] : []
}

function addVisualPanel(slide, slideData, assets, box, options = {}) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    rectRadius: 0.1,
    fill: { color: options.fill || palette.paper, transparency: options.transparency ?? 2 },
    line: { color: options.line || palette.line },
  })
  const usable = assets.slice(0, options.max || 4)
  if (!usable.length) {
    slide.addText(safeText(slideData.imagePlan, '시각 자료'), {
      x: box.x + 0.25,
      y: box.y + box.h / 2 - 0.25,
      w: box.w - 0.5,
      h: 0.5,
      fontSize: 13,
      color: palette.muted,
      align: 'center',
      fit: 'shrink',
      margin: 0.02,
    })
    return
  }
  if (usable.length === 1) {
    addAsset(slide, usable[0], { x: box.x + 0.45, y: box.y + 0.38, w: box.w - 0.9, h: box.h - 0.85 }, { label: options.label })
    return
  }
  const cols = usable.length <= 2 ? usable.length : 2
  const rows = Math.ceil(usable.length / cols)
  const gap = 0.22
  const cellW = (box.w - 0.7 - gap * (cols - 1)) / cols
  const cellH = (box.h - 0.62 - gap * (rows - 1)) / rows
  usable.forEach((asset, index) => {
    const col = index % cols
    const row = Math.floor(index / cols)
    addAsset(slide, asset, {
      x: box.x + 0.35 + col * (cellW + gap),
      y: box.y + 0.28 + row * (cellH + gap),
      w: cellW,
      h: cellH - 0.17,
    }, { label: options.label !== false, labelSize: 6.5 })
  })
}

function addTextChips(slide, items, boxes, options = {}) {
  boxes.forEach((box, index) => {
    const text = items[index] || ''
    if (!text) return
    slide.addShape(pptx.ShapeType.roundRect, {
      x: box.x,
      y: box.y,
      w: box.w,
      h: box.h,
      rectRadius: 0.06,
      fill: { color: box.fill || options.fill || 'FFFFFF' },
      line: { color: box.line || options.line || palette.line },
    })
    slide.addText(text, {
      x: box.x + 0.18,
      y: box.y + 0.14,
      w: box.w - 0.36,
      h: box.h - 0.22,
      fontSize: box.fontSize || options.fontSize || 13,
      bold: box.bold ?? options.bold ?? true,
      color: box.color || palette.ink,
      fit: 'shrink',
      margin: 0.02,
    })
  })
}

function addTeacherNote(slide, slideData, box, dark = false) {
  const text = `도식: ${safeText(slideData.diagramPlan, '없음')}\n교사용 메모: ${safeText(slideData.teacherNote)}`
  slide.addText(text, {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    fontSize: 8.5,
    color: dark ? 'CFE3C4' : '50695E',
    fit: 'shrink',
    margin: 0.02,
  })
}

function coverLayout(slide, slideData) {
  addTopBand(slide, slideData, { dark: true, titleW: 6.6, titleSize: 34 })
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.55,
    y: 2.35,
    w: 5.9,
    h: 1.25,
    rectRadius: 0.08,
    fill: { color: '214A3B' },
    line: { color: '3C6B55' },
  })
  addMainMessage(slide, slideData, { x: 0.85, y: 2.58, w: 5.3, h: 0.72, fontSize: 18, bold: true, color: 'FFFFFF' })
  addActivityCard(slide, slideData, { x: 0.72, y: 4.15, w: 5.4, h: 1.15 }, { fill: 'F5E8BD', line: 'CFAF6B', accent: palette.accent })
  const assets = pickAssets(slideData, 5)
  addVisualPanel(slide, slideData, assets, { x: 7.0, y: 0.82, w: 5.7, h: 5.85 }, { fill: 'F6F0DD', max: 5 })
  slide.addText('SITPO · 에셋 기반 수업 슬라이드', {
    x: 0.7,
    y: 6.42,
    w: 5.2,
    h: 0.28,
    fontSize: 11,
    color: 'CFE3C4',
    margin: 0,
  })
}

function visualLeftLayout(slide, slideData) {
  addTopBand(slide, slideData, { titleW: 7.2 })
  const assets = pickAssets(slideData, 3)
  addVisualPanel(slide, slideData, assets, { x: 0.65, y: 1.65, w: 5.25, h: 4.65 }, { max: 3, fill: 'FFF9EA' })
  addMainMessage(slide, slideData, { x: 6.35, y: 1.62, w: 5.85, h: 0.72, fontSize: 18, bold: true })
  addTextChips(slide, visibleText(slideData), [
    { x: 6.35, y: 2.65, w: 2.7, h: 0.95, fill: 'E6F1D2' },
    { x: 9.25, y: 2.65, w: 2.7, h: 0.95, fill: 'F4E4C1' },
    { x: 6.35, y: 3.85, w: 2.7, h: 0.95, fill: 'E4EEF3' },
    { x: 9.25, y: 3.85, w: 2.7, h: 0.95, fill: 'F5DED8' },
  ], { fontSize: 12.5 })
  addActivityCard(slide, slideData, { x: 6.35, y: 5.28, w: 5.6, h: 0.92 })
  addTeacherNote(slide, slideData, { x: 0.65, y: 6.55, w: 10.5, h: 0.35 })
}

function visualRightLayout(slide, slideData) {
  addTopBand(slide, slideData, { titleW: 8.2 })
  addMainMessage(slide, slideData, { x: 0.65, y: 1.55, w: 5.15, h: 0.78, fontSize: 18, bold: true })
  const items = visibleText(slideData)
  items.slice(0, 4).forEach((text, index) => {
    const y = 2.65 + index * 0.72
    slide.addShape(pptx.ShapeType.ellipse, {
      x: 0.72,
      y: y + 0.02,
      w: 0.32,
      h: 0.32,
      fill: { color: [palette.green, palette.accent, palette.blue, palette.coral][index % 4] },
      line: { color: 'FFFFFF' },
    })
    slide.addText(text, {
      x: 1.18,
      y,
      w: 4.8,
      h: 0.45,
      fontSize: 14,
      bold: true,
      color: palette.ink,
      fit: 'shrink',
      margin: 0,
    })
  })
  addActivityCard(slide, slideData, { x: 0.65, y: 5.55, w: 5.4, h: 0.9 }, { fill: 'F4E4C1' })
  addVisualPanel(slide, slideData, pickAssets(slideData, 4), { x: 6.55, y: 1.45, w: 5.8, h: 4.95 }, { max: 4, fill: 'FFFFFF' })
  addTeacherNote(slide, slideData, { x: 6.7, y: 6.55, w: 5.2, h: 0.32 })
}

function cardGridLayout(slide, slideData) {
  addTopBand(slide, slideData, { titleW: 8.5 })
  addMainMessage(slide, slideData, { x: 0.65, y: 1.45, w: 11.7, h: 0.48, fontSize: 16, bold: true })
  const assets = pickAssets(slideData, 4)
  const cards = [
    { x: 0.72, y: 2.18, w: 2.9, h: 2.05, fill: 'FFFDF4' },
    { x: 3.92, y: 2.18, w: 2.9, h: 2.05, fill: 'E6F1D2' },
    { x: 7.12, y: 2.18, w: 2.9, h: 2.05, fill: 'F4E4C1' },
    { x: 10.32, y: 2.18, w: 2.25, h: 2.05, fill: 'E4EEF3' },
  ]
  const texts = visibleText(slideData)
  const count = Math.min(cards.length, Math.max(texts.length, assets.length, 1))
  cards.slice(0, count).forEach((box, index) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: box.x,
      y: box.y,
      w: box.w,
      h: box.h,
      rectRadius: 0.08,
      fill: { color: box.fill },
      line: { color: palette.line },
    })
    if (assets[index]) addAsset(slide, assets[index], { x: box.x + 0.35, y: box.y + 0.16, w: box.w - 0.7, h: 0.9 }, { label: false })
    slide.addText(texts[index] || safeText(assets[index]?.name, safeText(slideData.diagramPlan, '정리하기')), {
      x: box.x + 0.2,
      y: box.y + 1.2,
      w: box.w - 0.4,
      h: 0.55,
      fontSize: 12.5,
      bold: true,
      color: palette.ink,
      align: 'center',
      fit: 'shrink',
      margin: 0.02,
    })
  })
  addActivityCard(slide, slideData, { x: 0.72, y: 4.72, w: 5.7, h: 1.05 }, { fill: 'F5DED8' })
  addTeacherNote(slide, slideData, { x: 7.0, y: 4.8, w: 5.1, h: 0.9 })
}

function processLayout(slide, slideData) {
  addTopBand(slide, slideData, { dark: true, titleW: 8.1 })
  addMainMessage(slide, slideData, { x: 0.65, y: 1.55, w: 11.5, h: 0.55, fontSize: 17, bold: true, color: 'FFFFFF' })
  const items = visibleText(slideData).slice(0, 4)
  const assets = pickAssets(slideData, 4)
  const xs = [0.85, 3.85, 6.85, 9.85]
  const count = Math.min(xs.length, Math.max(items.length, assets.length, 1))
  xs.slice(0, count).forEach((x, index) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y: 2.55 + (index % 2) * 0.42,
      w: 2.18,
      h: 1.9,
      rectRadius: 0.08,
      fill: { color: ['F6F0DD', 'E6F1D2', 'F4E4C1', 'E4EEF3'][index % 4] },
      line: { color: '8FB56A' },
    })
    if (assets[index]) addAsset(slide, assets[index], { x: x + 0.35, y: 2.72 + (index % 2) * 0.42, w: 1.45, h: 0.82 }, { label: false })
    slide.addText(items[index] || safeText(assets[index]?.name, `활동 ${index + 1}`), {
      x: x + 0.18,
      y: 3.68 + (index % 2) * 0.42,
      w: 1.82,
      h: 0.45,
      fontSize: 11.5,
      bold: true,
      color: palette.ink,
      align: 'center',
      fit: 'shrink',
      margin: 0.02,
    })
    if (index < count - 1) {
      slide.addText('→', {
        x: x + 2.38,
        y: 3.15,
        w: 0.35,
        h: 0.28,
        fontSize: 20,
        bold: true,
        color: 'F5E8BD',
        margin: 0,
      })
    }
  })
  addActivityCard(slide, slideData, { x: 1.1, y: 5.65, w: 10.8, h: 0.75 }, { fill: '214A3B', line: '3C6B55', accent: 'F5E8BD', textColor: 'FFFFFF', fontSize: 10 })
  addTeacherNote(slide, slideData, { x: 0.75, y: 6.65, w: 10.8, h: 0.32 }, true)
}

function comparisonLayout(slide, slideData) {
  addTopBand(slide, slideData, { titleW: 8.2 })
  addMainMessage(slide, slideData, { x: 0.65, y: 1.45, w: 11.5, h: 0.45, fontSize: 16, bold: true })
  const groups = splitItems(visibleText(slideData), 2)
  const assets = pickAssets(slideData, 4)
  const columns = [
    { x: 0.75, title: '먼저 살펴보기', fill: 'E6F1D2', accent: palette.green, assets: assets.slice(0, 2), items: groups[0] },
    { x: 6.85, title: '비교하며 정리하기', fill: 'F4E4C1', accent: palette.accent, assets: assets.slice(2, 4), items: groups[1] },
  ]
  columns.forEach((col) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: col.x,
      y: 2.18,
      w: 5.65,
      h: 3.65,
      rectRadius: 0.08,
      fill: { color: col.fill },
      line: { color: palette.line },
    })
    slide.addText(col.title, {
      x: col.x + 0.28,
      y: 2.38,
      w: 4.9,
      h: 0.25,
      fontSize: 14,
      bold: true,
      color: col.accent,
      margin: 0,
    })
    addVisualPanel(slide, slideData, col.assets, { x: col.x + 0.25, y: 2.82, w: 2.4, h: 1.85 }, { max: 2, fill: 'FFFFFF', label: false })
    slide.addText((col.items.length ? col.items : ['차이점을 말해 보기']).map((text) => `• ${text}`).join('\n'), {
      x: col.x + 2.95,
      y: 2.95,
      w: 2.25,
      h: 1.45,
      fontSize: 11.5,
      bold: true,
      color: palette.ink,
      fit: 'shrink',
      margin: 0.02,
    })
  })
  addActivityCard(slide, slideData, { x: 1.1, y: 6.1, w: 10.8, h: 0.62 }, { fill: 'FFFDF4', fontSize: 10 })
}

function activityLayout(slide, slideData) {
  addTopBand(slide, slideData, { titleW: 8.4 })
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.72,
    y: 1.6,
    w: 7.35,
    h: 4.75,
    rectRadius: 0.1,
    fill: { color: 'FFFDF4' },
    line: { color: palette.line },
  })
  slide.addText('학생 활동 중심', {
    x: 1.08,
    y: 1.9,
    w: 3.5,
    h: 0.28,
    fontSize: 13,
    bold: true,
    color: palette.accent,
    margin: 0,
  })
  addMainMessage(slide, slideData, { x: 1.08, y: 2.28, w: 6.55, h: 0.62, fontSize: 19, bold: true })
  slide.addText(visibleText(slideData).map((text, index) => `${index + 1}. ${text}`).join('\n'), {
    x: 1.1,
    y: 3.2,
    w: 3.55,
    h: 1.35,
    fontSize: 13.5,
    bold: true,
    color: palette.ink,
    fit: 'shrink',
    margin: 0.02,
  })
  addActivityCard(slide, slideData, { x: 1.05, y: 4.95, w: 6.45, h: 0.95 }, { fill: 'E6F1D2', fontSize: 12 })
  const assets = pickAssets(slideData, 3)
  addVisualPanel(slide, slideData, assets, { x: 8.5, y: 1.55, w: 3.55, h: 4.8 }, { max: 3, fill: 'F4E4C1' })
  addTeacherNote(slide, slideData, { x: 0.9, y: 6.58, w: 10.8, h: 0.35 })
}

function summaryLayout(slide, slideData) {
  addTopBand(slide, slideData, { dark: true, titleW: 8.8, titleSize: 30 })
  const assets = pickAssets(slideData, 6)
  addMainMessage(slide, slideData, { x: 0.7, y: 1.65, w: 6.2, h: 0.7, fontSize: 19, bold: true, color: 'FFFFFF' })
  addTextChips(slide, visibleText(slideData), [
    { x: 0.75, y: 2.75, w: 2.8, h: 0.82, fill: '214A3B', line: '3C6B55', color: 'FFFFFF' },
    { x: 3.85, y: 2.75, w: 2.8, h: 0.82, fill: '214A3B', line: '3C6B55', color: 'FFFFFF' },
    { x: 0.75, y: 3.9, w: 2.8, h: 0.82, fill: '214A3B', line: '3C6B55', color: 'FFFFFF' },
    { x: 3.85, y: 3.9, w: 2.8, h: 0.82, fill: '214A3B', line: '3C6B55', color: 'FFFFFF' },
  ], { fontSize: 12.5 })
  addVisualPanel(slide, slideData, assets, { x: 7.25, y: 1.12, w: 5.25, h: 5.35 }, { max: 6, fill: 'F6F0DD' })
  addActivityCard(slide, slideData, { x: 0.75, y: 5.52, w: 5.9, h: 0.78 }, { fill: 'F5E8BD', line: 'CFAF6B', accent: palette.accent, fontSize: 10 })
}

function layoutFor(slideData, index, total) {
  if (index === 0) return coverLayout
  if (index === total - 1) return summaryLayout
  const cycle = [visualLeftLayout, cardGridLayout, processLayout, visualRightLayout, comparisonLayout, activityLayout]
  return cycle[(index - 1) % cycle.length]
}

const pptx = new PptxGenJS()
pptx.layout = 'LAYOUT_WIDE'
pptx.author = 'SITPO Codex Job API'
pptx.subject = `${project.grade} ${project.subject} ${project.unit}`
pptx.title = project.title
pptx.company = 'SITPO'
pptx.lang = 'ko-KR'
pptx.theme = {
  headFontFace: 'Arial',
  bodyFontFace: 'Arial',
  lang: 'ko-KR',
}
pptx.defineLayout({ name: 'LAYOUT_WIDE', width: W, height: H })

project.slides.forEach((slideData, index) => {
  const slide = pptx.addSlide()
  slide.background = { color: palette.cream }
  const layout = layoutFor(slideData, index, project.slides.length)
  layout(slide, slideData, index)
  addPageNumber(slide, slideData.slideNo || index + 1)
})

fs.mkdirSync(path.dirname(outputPath), { recursive: true })
await pptx.writeFile({ fileName: outputPath })
console.log(`Wrote ${outputPath}`)
