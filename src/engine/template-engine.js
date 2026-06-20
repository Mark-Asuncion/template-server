const { DocxTemplate } = require('docx-template')
const JSZip = require('jszip')
const saveAs = require('file-saver')
const fs = require("fs");
const path = require("path");

const WORD_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const TEMPLATE_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

const PLACEHOLDER_REPLACEMENTS = [
  { placeholder: '[Insert Date]', expression: 'date', occurrence: 0 },
  { placeholder: '[Insert Date]', expression: 'dateOfLoss', occurrence: 1 },
  { placeholder: '[Insert Client Name]', expression: 'claimantName' },
  { placeholder: '[Insert Client Address]', expression: 'propertyAddress' },
  { placeholder: '[Insert Title/Position]', expression: 'preparedByTitle' },
  { placeholder: '[Insert]', expression: 'lossType' },
  { placeholder: '[Insert Insured Name]', expression: 'insuredName' },
  { placeholder: '[Insert Location]', expression: 'locationOfRisk' },
  { placeholder: '[Insert Policy Number]', expression: 'claimNumber' },
  { placeholder: '[Insert Principal/Claims Department/Loss Manager Name]', expression: 'contactPersonName' },
  { placeholder: '[Typhoon: Include brief description/discussion]', expression: 'typhoonDescription' },
  { placeholder: '[Insert Name]', expression: 'typhoonName', occurrence: 0 },
  { placeholder: '[Insert km/h or mph]', expression: 'windSpeed' },
  { placeholder: '[Insert Wind Signal No., e.g., Signal No. 3]', expression: 'windSignal' },
  { placeholder: '[Flood: Include brief description/discussion]', expression: 'floodDescription' },
  { placeholder: '[Insert flood level, e.g., 1.5 meters]', expression: 'floodLevel' },
  { placeholder: '[Date Assignment Received]', expression: 'assignmentDate' },
  { placeholder: '[Name of Insured]', expression: 'insuredName' },
  { placeholder: '[Insurer Name]', expression: 'insurerName' },
  { placeholder: '[Location of Risk]', expression: 'locationOfRisk' },
  { placeholder: '[Date of Loss]', expression: 'dateOfLoss' },
  { placeholder: '[Date of Site Attendance]', expression: 'siteAttendanceDate' },
  { placeholder: '[Insert Property Address]', expression: 'propertyAddress' },
  { placeholder: '[Insert Contact Person or Representative’s Name]', expression: 'contactPersonName' },
  { placeholder: '[Insert Name]', expression: 'preparedByName', occurrence: 1 },
  { placeholder: '[Insert Title/Position]', expression: 'preparedByTitle', occurrence: 1 }
]

function toText(value) {
  if (value === undefined || value === null) return ''
  if (Array.isArray(value)) return value.map(toText).filter(Boolean).join(', ')
  if (typeof value === 'object') return ''
  return String(value).replace(/\s+/g, ' ').trim()
}

function firstDefined(...values) {
  return values.find((value) => value !== undefined && value !== null && value !== '') || ''
}

function buildTemplateData(payload = {}) {
  const workflowMetadata = payload.workflowMetadata || {}
  const inputDates = workflowMetadata.inputDates || {}
  const claimantName = toText(firstDefined(payload.claimantName, payload.claimant, payload.insuredName, payload.insured))
  const insurer = toText(firstDefined(payload.insurer, payload.insuranceCompany, payload.insurance_company))
  const locationOfRisk = toText(firstDefined(
    payload.locationOfRisk,
    payload.location_of_risk,
    payload.propertyAddress,
    payload.property_address
  ))
  const summary = toText(firstDefined(payload.summary, payload.reportSummary, payload.report_summary, payload.analysis))
  const recommendation = toText(firstDefined(payload.recommendation, payload.findings, payload.nextSteps, payload.next_steps))
  const date = toText(firstDefined(payload.date, payload.reportDate, payload.report_date, new Date().toLocaleDateString()))

  return {
    ...payload,
    title: toText(firstDefined(payload.title, payload.reportTitle, payload.report_title, 'Claim Assistance Report')),
    claimNumber: toText(firstDefined(payload.claimNumber, payload.claim_number, payload.claimId, payload.claim_id, 'N/A')),
    claimantName,
    insuredName: toText(firstDefined(payload.insuredName, payload.insured_name, payload.insured, claimantName)),
    insurer,
    insurerName: toText(firstDefined(payload.insurerName, payload.insurer_name, insurer)),
    locationOfRisk,
    propertyAddress: toText(firstDefined(payload.propertyAddress, payload.property_address, locationOfRisk)),
    date,
    dateOfLoss: toText(firstDefined(payload.dateOfLoss, payload.date_of_loss, payload.lossDate, payload.loss_date, date)),
    assignmentDate: toText(firstDefined(payload.assignmentDate, payload.assignment_date, inputDates.documentsReceivedAt)),
    siteAttendanceDate: toText(firstDefined(payload.siteAttendanceDate, payload.site_attendance_date, inputDates.inspectionCompletedDate)),
    lossType: toText(firstDefined(payload.lossType, payload.loss_type, payload.peril, payload.causeOfLoss, payload.cause_of_loss, 'typhoon/flood')),
    typhoonDescription: toText(firstDefined(
      payload.typhoonDescription,
      payload.typhoon_description,
      payload.eventDescription,
      payload.event_description,
      summary
    )),
    typhoonName: toText(firstDefined(payload.typhoonName, payload.typhoon_name, payload.stormName, payload.storm_name)),
    windSpeed: toText(firstDefined(payload.windSpeed, payload.wind_speed, payload.windVelocity, payload.wind_velocity)),
    windSignal: toText(firstDefined(payload.windSignal, payload.wind_signal, payload.windSignalNo, payload.wind_signal_no)),
    floodDescription: toText(firstDefined(
      payload.floodDescription,
      payload.flood_description,
      payload.floodingDescription,
      payload.flooding_description,
      summary
    )),
    floodLevel: toText(firstDefined(payload.floodLevel, payload.flood_level, payload.floodHeight, payload.flood_height)),
    contactPersonName: toText(firstDefined(
      payload.contactPersonName,
      payload.contact_person_name,
      payload.representativeName,
      payload.representative_name
    )),
    preparedByName: toText(firstDefined(payload.preparedByName, payload.prepared_by_name, payload.authorName, payload.author_name, 'NOVA')),
    preparedByTitle: toText(firstDefined(payload.preparedByTitle, payload.prepared_by_title, 'Claims Assistant')),
    summary,
    recommendation,
    workflowMetadata,
    attachments: Array.isArray(payload.attachments) ? payload.attachments : [],
    images: Array.isArray(payload.images) ? payload.images : []
  }
}

function getTextNodes(document) {
  return Array.from(document.getElementsByTagNameNS(WORD_NAMESPACE, 't'))
}

function replaceTextRange(textNodes, startIndex, endIndex, replacement) {
  let offset = 0
  let startNodeIndex = -1
  let startOffset = 0
  let endNodeIndex = -1
  let endOffset = 0

  for (let index = 0; index < textNodes.length; index += 1) {
    const nodeText = textNodes[index].textContent || ''
    const nextOffset = offset + nodeText.length

    if (startNodeIndex === -1 && startIndex < nextOffset) {
      startNodeIndex = index
      startOffset = startIndex - offset
    }

    if (endIndex <= nextOffset) {
      endNodeIndex = index
      endOffset = endIndex - offset
      break
    }

    offset = nextOffset
  }

  if (startNodeIndex === -1 || endNodeIndex === -1) return

  const startNode = textNodes[startNodeIndex]
  const endNode = textNodes[endNodeIndex]
  const prefix = startNode.textContent.slice(0, startOffset)
  const suffix = endNode.textContent.slice(endOffset)

  startNode.textContent = replacement

  for (let index = startNodeIndex + 1; index < endNodeIndex; index += 1) {
    textNodes[index].textContent = ''
  }

  if (endNodeIndex !== startNodeIndex) {
    endNode.textContent = suffix
  } else {
    startNode.textContent = `${prefix}${replacement}${suffix}`
  }
}

function replacePlaceholders(xml) {
  const parser = new DOMParser()
  const document = parser.parseFromString(xml, 'application/xml')
  const parseError = document.querySelector('parsererror')

  if (parseError) {
    throw new Error('The report template XML could not be parsed.')
  }

  const textNodes = getTextNodes(document)
  const combinedText = textNodes.map((node) => node.textContent || '').join('')
  const replacements = []

  PLACEHOLDER_REPLACEMENTS.forEach((replacement) => {
    const occurrenceCounts = new Map()
    let searchFrom = 0

    while (searchFrom < combinedText.length) {
      const startIndex = combinedText.indexOf(replacement.placeholder, searchFrom)
      if (startIndex === -1) break

      const count = occurrenceCounts.get(replacement.placeholder) || 0
      occurrenceCounts.set(replacement.placeholder, count + 1)

      if (replacement.occurrence === undefined || replacement.occurrence === count) {
        replacements.push({
          startIndex,
          endIndex: startIndex + replacement.placeholder.length,
          expression: replacement.expression
        })
      }

      searchFrom = startIndex + 1
    }
  })

  replacements.sort((left, right) => right.startIndex - left.startIndex).forEach((replacement) => {
    replaceTextRange(textNodes, replacement.startIndex, replacement.endIndex, `\${${replacement.expression}}`)
  })

  const remainingPlaceholders = Array.from(new Set(
    textNodes.map((node) => node.textContent || '').join('').match(/\[[^\]]+\]/g) || []
  ))

  if (remainingPlaceholders.length > 0) {
    throw new Error(`Report template still has unmapped placeholders: ${remainingPlaceholders.join(', ')}`)
  }

  return new XMLSerializer().serializeToString(document)
}

async function loadTemplateBlob() {
    // static for now
    const filePath = path.join(process.cwd(), "public", "docx", "TCS_Advance-Report_Typhoon_Flood_Template_-20250730.docx");
    const arrayBuffer = [];
    fs.readFile(filePath, "utf8", (err, data) => {
        if (err) {
            throw new Error("failed to read the path: ", path)
        }
        arrayBuffer = data.buffer.slice(
            data.byteOffset,
            data.byteOffset + data.byteLength
        );
    });

    if (arrayBuffer.length == 0) {
        throw new Error('The report template could not be loaded.')
    }

    const templateArrayBuffer = arrayBuffer;
    const sourceZip = new JSZip(templateArrayBuffer)
    const outputZip = new JSZip()

    Object.keys(sourceZip.files).forEach((fileName) => {
        const file = sourceZip.files[fileName]

        if (file.dir) {
            outputZip.file(fileName, null, { dir: true })
            return
        }

        if (fileName === 'word/document.xml') {
            outputZip.file(fileName, replacePlaceholders(file.asText()))
            return
        }

        outputZip.file(fileName, file.asBinary(), { binary: true })
    })

    return outputZip.generate({ type: 'blob', mimeType: TEMPLATE_MIME_TYPE })
}

function useDocxReportGenerator() {
    const generate = async (payload, fileName = `NOVA-report-${Date.now()}.docx`) => {
        if (typeof window === 'undefined') {
            throw new Error('Document generation requires a browser environment.')
        }

        const templateBlob = await loadTemplateBlob()
        const templateData = buildTemplateData(payload)
        const doc = await DocxTemplate.assemble(templateBlob, templateData, { clearWrap: true })
        const blob = doc.serialize().generate({ type: 'blob', mimeType: TEMPLATE_MIME_TYPE })

        const filePath = path.join(process.cwd(), "public", "generated", fileName);

        saveAs(blob, filePath)

        return path.join("public", "generated", fileName);
    }

    return { generate }
}

module.exports = useDocxReportGenerator;
