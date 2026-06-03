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

for (const slideData of project.slides) {
  const slide = pptx.addSlide()
  slide.background = { color: 'F8F2E4' }

  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.35,
    fill: { color: '12392D' },
    line: { color: '12392D' },
  })

  slide.addText(`${project.grade} ${project.subject} · ${project.unit}`, {
    x: 0.55,
    y: 0.55,
    w: 8.8,
    h: 0.25,
    fontSize: 9,
    bold: true,
    color: '537466',
    margin: 0,
  })

  slide.addText(slideData.title, {
    x: 0.55,
    y: 0.85,
    w: 8.65,
    h: 0.62,
    fontSize: 25,
    bold: true,
    color: '16352C',
    fit: 'shrink',
    margin: 0.02,
  })

  slide.addText(slideData.mainMessage, {
    x: 0.58,
    y: 1.6,
    w: 8.35,
    h: 0.65,
    fontSize: 15,
    color: '385A4C',
    fit: 'shrink',
    breakLine: false,
    margin: 0.02,
  })

  const visibleText = Array.isArray(slideData.visibleText) ? slideData.visibleText : []
  slide.addText(visibleText.map((text) => `• ${text}`).join('\n'), {
    x: 0.68,
    y: 2.35,
    w: 4.6,
    h: 1.3,
    fontSize: 16,
    bold: true,
    color: '17352C',
    breakLine: false,
    fit: 'shrink',
    margin: 0.04,
  })

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.55,
    y: 4.2,
    w: 5.2,
    h: 1,
    rectRadius: 0.08,
    fill: { color: 'E6F1D2' },
    line: { color: 'CAD9BE' },
  })
  slide.addText(`활동\n${slideData.studentActivity}`, {
    x: 0.75,
    y: 4.34,
    w: 4.8,
    h: 0.72,
    fontSize: 12,
    color: '214537',
    fit: 'shrink',
    margin: 0.02,
  })

  const slideAssets = assetsForSlide(slideData)
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 6.05,
    y: 1.02,
    w: 6.38,
    h: 3.95,
    rectRadius: 0.08,
    fill: { color: 'FFFDF4', transparency: 4 },
    line: { color: 'D8E1D3' },
  })

  if (slideAssets.length > 0) {
    const slots = [
      { x: 6.35, y: 1.22, w: 2.75, h: 1.55 },
      { x: 9.25, y: 1.22, w: 2.75, h: 1.55 },
      { x: 6.35, y: 2.98, w: 2.75, h: 1.55 },
      { x: 9.25, y: 2.98, w: 2.75, h: 1.55 },
    ]
    slideAssets.slice(0, 4).forEach((asset, index) => {
      const slot = slots[index]
      slide.addImage({ path: asset.imagePath, x: slot.x, y: slot.y, w: slot.w, h: slot.h, sizingContain: true })
      slide.addText(asset.name || asset.id || '', {
        x: slot.x,
        y: slot.y + slot.h + 0.03,
        w: slot.w,
        h: 0.18,
        fontSize: 7,
        color: '617568',
        align: 'center',
        margin: 0,
        fit: 'shrink',
      })
    })
  } else {
    const sheet = sheetImagePath()
    if (sheet) {
      slide.addImage({ path: sheet, x: 6.25, y: 1.2, w: 5.95, h: 3.55, sizingContain: true })
    } else {
      slide.addText(slideData.imagePlan || '이미지 계획 없음', {
        x: 6.45,
        y: 2.45,
        w: 5.75,
        h: 0.9,
        fontSize: 13,
        color: '617568',
        align: 'center',
        fit: 'shrink',
        margin: 0.02,
      })
    }
  }

  slide.addText(`도식: ${slideData.diagramPlan || '없음'}\n교사용 메모: ${slideData.teacherNote || ''}`, {
    x: 6.22,
    y: 5.08,
    w: 6.22,
    h: 0.65,
    fontSize: 10,
    color: '50695E',
    fit: 'shrink',
    margin: 0.02,
  })

  slide.addText(String(slideData.slideNo).padStart(2, '0'), {
    x: 12.26,
    y: 6.78,
    w: 0.7,
    h: 0.22,
    fontSize: 8,
    bold: true,
    color: '6C7E74',
    align: 'right',
    margin: 0,
  })
}

fs.mkdirSync(path.dirname(outputPath), { recursive: true })
await pptx.writeFile({ fileName: outputPath })
console.log(`Wrote ${outputPath}`)
