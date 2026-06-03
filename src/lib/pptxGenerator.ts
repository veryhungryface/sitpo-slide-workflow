import PptxGenJS from 'pptxgenjs'
import type { SitpoProject } from '../types'

export async function generatePptx(project: SitpoProject): Promise<Blob> {
  const pptx = new PptxGenJS()
  pptx.layout = 'LAYOUT_WIDE' // 16:9

  for (const slideData of project.slides) {
    const slide = pptx.addSlide()

    // Title
    slide.addText(slideData.title, {
      x: 0.5, y: 0.5, w: 9, h: 0.75,
      fontSize: 28, bold: true, color: '17352c',
      align: 'left', valign: 'middle',
    })

    // Main Message
    slide.addText(slideData.mainMessage, {
      x: 0.5, y: 1.5, w: 9, h: 0.5,
      fontSize: 18, color: '4f645a',
      align: 'left', valign: 'top',
    })

    // Visible Text (as bullet points or simple text blocks)
    if (slideData.visibleText.length > 0) {
      slide.addText(slideData.visibleText.map(t => `- ${t}`).join('\n'), {
        x: 0.5, y: 2.2, w: 9, h: 1.5,
        fontSize: 16, color: '17352c',
        align: 'left', valign: 'top',
        bullet: true,
      })
    }

    // Student Activity (simplified as a text block)
    if (slideData.studentActivity) {
      slide.addText(`활동: ${slideData.studentActivity}`, {
        x: 0.5, y: 4, w: 9, h: 0.5,
        fontSize: 14, color: '577164', italic: true,
        align: 'left', valign: 'top',
      })
    }

    // Image/Diagram Plan (simplified for now)
    slide.addText(`이미지 계획: ${slideData.imagePlan || '없음'}\n도식 계획: ${slideData.diagramPlan || '없음'}`, {
      x: 0.5, y: 5, w: 9, h: 0.75,
      fontSize: 12, color: '667a70',
      align: 'left', valign: 'top',
    })
  }

  const pptxContent = await pptx.write({ outputType: 'blob' })
  const mimeType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'

  if (pptxContent instanceof Blob) {
    return pptxContent
  }

  if (pptxContent instanceof ArrayBuffer) {
    return new Blob([pptxContent], { type: mimeType })
  }

  if (pptxContent instanceof Uint8Array) {
    const bytes = new Uint8Array(pptxContent.byteLength)
    bytes.set(pptxContent)
    return new Blob([bytes], { type: mimeType })
  }

  return new Blob([pptxContent], { type: mimeType })
}
