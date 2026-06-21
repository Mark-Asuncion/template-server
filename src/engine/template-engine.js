const fs = require('fs');
const path = require('path');
const PDFDocument = require('pdfkit');

function buildTemplateData(payload = {}) {
    const workflowMetadata = payload.workflowMetadata || {};
    const inputDates = workflowMetadata.inputDates || {};
    const firstDefined = (...vals) => vals.find(v => v !== undefined && v !== null && v !== '') || '';
    const toText = (val) => (typeof val === 'object' ? '' : String(val).trim());

    return {
        referenceNumber: toText(firstDefined(payload.referenceNumber, "________")),
        claimNumber: toText(firstDefined(payload.claimNumber, 'Please Advise')),
        date: toText(firstDefined(payload.date, '[Insert Date]')),
        insurer: toText(firstDefined(payload.insurer, "[Insert Client Name]")),
        clientAddress: toText(firstDefined(payload.clientAddress, "[Insert Client Address]")),
        attentionPerson: toText(firstDefined(payload.attentionPerson, "__________________")),
        attentionTitle: toText(firstDefined(payload.attentionTitle, "[Insert Title/Position]")),
        lossType: toText(firstDefined(payload.lossType, "[Insert]")),
        dateOfLoss: toText(firstDefined(payload.dateOfLoss, "[Insert Date]")),
        insuredName: toText(firstDefined(payload.claimantName, payload.insuredName, "[Insert Insured Name]")),
        locationOfRisk: toText(firstDefined(payload.locationOfRisk, "[Insert Location]")),
        policyNumber: toText(firstDefined(payload.policyNumber, "[Insert Policy Number]")),
        claimsManager: toText(firstDefined(payload.claimsManager, "[Insert Principal/Claims Department/Loss Manager Name]")),
        contactPerson: toText(firstDefined(payload.contactPerson, "[Insert Contact Person or Representative’s Name]")),
        siteAttendanceDate: toText(firstDefined(payload.siteAttendanceDate, inputDates.inspectionCompletedDate, "[Insert Date of Site Attendance]")),
        assignmentDate: toText(firstDefined(payload.assignmentDate, inputDates.documentsReceivedAt, "[Date Assignment Received]")),
        typhoonName: toText(firstDefined(payload.typhoonName, "[Insert Name]")),
        windSpeed: toText(firstDefined(payload.windSpeed, "[Insert km/h or mph]")),
        windSignal: toText(firstDefined(payload.windSignal, "[Insert Wind Signal No., e.g., Signal No. 3]")),
        floodLevel: toText(firstDefined(payload.floodLevel, "[Insert flood level, e.g., 1.5 meters]")),
        lossReserveAmount: toText(firstDefined(payload.lossReserveAmount, "___________")),
        preparedByName: toText(firstDefined(payload.preparedByName, '[Insert Name]')),
        preparedByTitle: toText(firstDefined(payload.preparedByTitle, '[Insert Title/Position]'))
    };
}

function useDocxReportGenerator() {
    const generate = async (payload, fileName = `NOVA-report-${Date.now()}.pdf`) => {
        const data = buildTemplateData(payload);
        const outputDir = path.join(process.cwd(), "public", "generated");
        if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });
        const outputPath = path.join(outputDir, fileName);

        return new Promise((resolve, reject) => {
            const doc = new PDFDocument({ size: 'A4', margin: 72 });
            const writeStream = fs.createWriteStream(outputPath);
            doc.pipe(writeStream);

            // --- 1. CORPORATE HEADER LETTERHEAD BLOCK ---
            const startY = 54;
            
            doc.fillColor('#000000').font('Helvetica-Bold').fontSize(11.5);
            doc.x = 72;
            doc.y = startY;
            doc.text("Total Claims Specialist, Inc.");
            
            doc.font('Helvetica', 8.5).fillColor('#333333');
            doc.y = startY + 15; doc.text("Room 403 & 405 Sikatuna Building., 6762 Ayala Avenue,");
            doc.y = startY + 25; doc.text("Makati City Philippines 1223");
            doc.y = startY + 35; doc.text("Telephone +63 2 5328 7070 Email: www.tcs.claims");
            doc.y = startY + 45; doc.text("GSM Globe # +63 9178193801 GSM Smart # +63 9188277137");

            const logoPath = path.join(process.cwd(), "public", "images", "tcs-logo.png");
            if (fs.existsSync(logoPath)) {
                doc.image(logoPath, 385, startY - 4, { width: 138 });
            }

            doc.x = 72;
            doc.y = startY + 75;
            doc.moveDown(1);

            // --- 2. REFERENCE METADATA AND ADDRESS BLOCKS ---
            doc.fillColor('#000000').font('Helvetica').fontSize(11);
            doc.text(`Our Reference No.: TCS/2025/${data.referenceNumber}`);
            doc.text(`Your Claim No.: ${data.claimNumber}`);
            doc.text(`${data.date}`);
            doc.moveDown(1.5);

            doc.text(data.insurer);
            doc.text(data.clientAddress);
            doc.moveDown(1);
            
            const tabLeftX = 72;
            const tabColonX = 145;
            const tabValueX = 160;

            let currentY = doc.y;
            doc.font('Helvetica');
            doc.x = tabLeftX; doc.y = currentY; doc.text("Attention");
            doc.x = tabColonX; doc.text(":");
            doc.x = tabValueX; doc.text(`MS./MR  ${data.attentionPerson}`);
            
            doc.x = tabValueX; doc.moveDown(0.1); doc.text(data.attentionTitle);
            doc.moveDown(2);

            // --- 3. SUBJECT DETAILS BLOCKS GRID MATRIX ---
            currentY = doc.y;
            doc.font('Helvetica-Bold');
            doc.x = tabLeftX; doc.y = currentY; doc.text("Subject");
            doc.x = tabColonX; doc.text(":");
            doc.x = tabValueX; doc.text("ADVANCE REPORT");
            doc.moveDown(0.8);

            const drawGridRow = (label, value) => {
                let rowY = doc.y;
                doc.font('Helvetica');
                doc.x = tabValueX; doc.y = rowY; doc.text(label, { width: 160 });
                doc.x = tabValueX + 165; doc.y = rowY; doc.text(":");
                doc.x = tabValueX + 175; doc.y = rowY; doc.text(value, { width: 175 });
                doc.moveDown(0.4);
            };

            drawGridRow("Claim for loss/damages due to", data.lossType);
            drawGridRow("Date of Loss", data.dateOfLoss);
            drawGridRow("Insured", data.insuredName);
            drawGridRow("Location of Risk", data.locationOfRisk);
            
            let policyY = doc.y;
            doc.font('Helvetica');
            doc.x = tabValueX + 110; doc.y = policyY; doc.text("Policy No.");
            doc.x = tabValueX + 165; doc.text(":");
            doc.x = tabValueX + 175; doc.text(data.policyNumber, { width: 175 });
            doc.moveDown(2);

            // --- 4. SALUTATION & PARAGRAPHS BLOCK ---
            doc.x = tabLeftX;
            doc.text(`Gentlemen/Ms./Mr. ${data.attentionPerson}:`);
            doc.moveDown(1);

            const printStandardParagraph = (text) => {
                doc.font('Helvetica').fontSize(11).text(text, { align: 'justify', lineGap: 3 });
                doc.moveDown(1.2);
            };

            printStandardParagraph(`We acknowledge with thanks the kind instruction of your ${data.claimsManager} for us to handle the investigation and adjustment of the subject typhoon-related loss/claim.`);
            printStandardParagraph(`In coordination with ${data.contactPerson}, the Insured, we confirm that we attended the Insured’s premises, located at ${data.locationOfRisk}, on ${data.siteAttendanceDate} for conducting an initial inspection and undertaking preliminary inquiries.`);
            printStandardParagraph(`Thereat, we met with the Insured, who cooperated fully and provided us with relevant information concerning the circumstances surrounding the incident and the extent and nature of the reported damage.`);
            printStandardParagraph(`We now present hereunder our preliminary findings for the Insurer’s detailed consideration.`);

            // --- 5. REPORT CONTENT CLAUSES ---
            doc.font('Helvetica-Bold').text('INSURED');
            doc.moveDown(0.5);
            printStandardParagraph(`In accordance with the assignment received on ${data.assignmentDate}, we acknowledge receipt of the claim filed. The claim pertains to the insured, ${data.insuredName} (“Insured”), who holds a policy with ${data.insurer} (“Insurer”), covering the property located at ${data.locationOfRisk}. The reported loss occurred on ${data.dateOfLoss}, coinciding with the passage of a typhoon. Preliminary information indicates that the loss resulted from wind and rain-related damage consistent with typhoon-related perils. Verification of the event through official meteorological data and local weather reports is underway to confirm the timing and nature of the storm event.`);

            doc.font('Helvetica-Bold').text('CIRCUMSTANCES/DISCOVERY OF LOSS');
            doc.moveDown(0.5);
            printStandardParagraph(`The Insured properties sustained damage due to flood/typhoon.`);
            printStandardParagraph(`[Typhoon: Include brief description/discussion] At the time of the loss, Typhoon ${data.typhoonName} affected the area with wind velocities reaching approximately ${data.windSpeed}, classified under ${data.windSignal} at the location of risk.`);
            printStandardParagraph(`[Flood: Include brief description/discussion] In conjunction with the typhoon, heavy rainfall contributed to significant surface flooding. Floodwaters reportedly rose to a height of ${data.floodLevel} within the insured premises, resulting in water ingress into ground-floor areas, damage to contents, equipment, and possible structural saturation.`);
            printStandardParagraph(`The reported cause of loss appears to be directly attributable to typhoon-force conditions.`);

            doc.font('Helvetica-Bold').text('CLAIM AND DOCUMENTATION');
            doc.moveDown(0.5);
            printStandardParagraph(`As of this writing, the Insured has yet to submit a formal statement of claim together with the requisite supporting documentation.`);
            printStandardParagraph(`We have issued written correspondence to the Insured requesting submission of the necessary documents to substantiate the claim. For the Insurer’s reference and file, enclosed is a copy of our letter addressed to the Insured outlining these requirements.`);
            printStandardParagraph(`In addition, we have reminded the Insured of their obligation under the Policy to take all reasonable measures to mitigate further loss or damage, including securing the premises and preserving undamaged property. We are monitoring compliance with this duty and will report further should any concerns arise in this regard.`);

            doc.font('Helvetica-Bold').text('LOSS RESERVE');
            doc.moveDown(0.5);
            printStandardParagraph(`Based on the initial data and observations available to date, we have provisionally estimated the cost of damages in the amount of PHP ${data.lossReserveAmount}, which the Insurers may consider setting aside as a precautionary loss reserve for this claim.`);
            printStandardParagraph(`Please note that this reserve is based on preliminary findings and may be subject to significant adjustment upon receipt of further, including the complete copy of the Policy, engineering assessments, contractor quotations, and supporting claim documentation.`);
            printStandardParagraph(`We shall continue to monitor, evaluate, and revise the suggested loss reserve as appropriate and shall promptly advise Insurers should any revision become necessary considering more definitive information.`);

            doc.font('Helvetica-Bold').text('PRESENT POSITION');
            doc.moveDown(0.5);
            printStandardParagraph(`The matter continues to receive our careful and ongoing attention. We shall provide further reports and updates as the claim develops and additional information becomes available.`);
            printStandardParagraph(`In the meantime, we would appreciate it if you could advise us should you have any specific comments, instructions, or directions, particularly with respect to the salvage aspect of the claim or any other area requiring immediate focus.`);
            printStandardParagraph(`We remain at your disposal for further guidance and will act accordingly upon receipt of your instructions.`);

            // --- 6. MULTI-COLUMN AUTHORIZED SIGN-OFFS BLOCK ---
            doc.moveDown(1.5);
            doc.font('Helvetica').text('Sincerely yours,');
            doc.moveDown(0.5);
            doc.font('Helvetica-Bold').text('For and on behalf of\nTOTAL CLAIMS SPECIALIST, INC.');
            doc.moveDown(4.5);

            const columnSigY = doc.y;
            
            doc.font('Helvetica-Bold');
            doc.x = 72; doc.y = columnSigY; doc.text(data.preparedByName);
            doc.font('Helvetica').text(data.preparedByTitle);

            doc.font('Helvetica-Bold');
            doc.x = 320; doc.y = columnSigY; doc.text('SHIELA L. DE DIOS');
            doc.font('Helvetica').text('Executive Vice President');
            doc.moveDown(3.5);

            doc.font('Helvetica-Bold');
            doc.x = 72; doc.text('VALENTINO G. ABOY');
            doc.font('Helvetica').text('President');

            doc.end();
            writeStream.on('finish', () => resolve(path.join("public", "generated", fileName)));
            writeStream.on('error', reject);
        });
    };

    return { generate };
}

module.exports = useDocxReportGenerator;