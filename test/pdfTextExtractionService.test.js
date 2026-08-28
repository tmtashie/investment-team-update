const test = require("node:test");
const assert = require("node:assert/strict");
const {
  MAX_PDF_UPLOAD_BYTES,
  validatePdfUpload,
  extractPdfTextFromUpload
} = require("../services/pdfTextExtractionService");

function makeTinyPdf(pages) {
  const streams = pages
    .map(
      (text, index) => `
${index + 1} 0 obj
<< /Length ${text.length + 20} >>
stream
BT /F1 12 Tf 72 720 Td (${text}) Tj ET
endstream
endobj`
    )
    .join("\n");
  return Buffer.from(`%PDF-1.4\n${streams}\n%%EOF`, "latin1");
}

function asUpload(buffer, filename = "FINSYNC July Update.pdf") {
  return {
    filename,
    mimeType: "application/pdf",
    fileData: buffer.toString("base64")
  };
}

test("extracts readable text from a PDF upload with page numbers", async () => {
  const result = await extractPdfTextFromUpload(
    asUpload(makeTinyPdf(["FINSYNC July revenue was $3.0 million.", "Customer count reached 62,824."]))
  );

  assert.equal(result.filename, "FINSYNC July Update.pdf");
  assert.equal(result.pages.length, 2);
  assert.equal(result.pages[0].pageNumber, 1);
  assert.match(result.combinedText, /Page 1:/);
  assert.match(result.combinedText, /FINSYNC July revenue/);
});

test("rejects non-PDF uploads before analysis", () => {
  assert.throws(
    () => validatePdfUpload(asUpload(Buffer.from("hello"), "update.txt")),
    /Only PDF files/
  );
});

test("rejects corrupt PDF uploads", () => {
  assert.throws(
    () => validatePdfUpload(asUpload(Buffer.from("not a pdf"), "update.pdf")),
    /not a valid PDF/
  );
});

test("rejects PDFs with no extractable text", async () => {
  await assert.rejects(
    () => extractPdfTextFromUpload(asUpload(Buffer.from("%PDF-1.4\n%%EOF", "latin1"))),
    /image-based|readable text/
  );
});

test("enforces PDF upload size limit", () => {
  const buffer = Buffer.concat([
    Buffer.from("%PDF-", "latin1"),
    Buffer.alloc(MAX_PDF_UPLOAD_BYTES + 1)
  ]);

  assert.throws(
    () => validatePdfUpload(asUpload(buffer, "large.pdf")),
    /limited to 10 MB/
  );
});

test("flags sparse financial pages for future OCR instead of metric extraction", async () => {
  const result = await extractPdfTextFromUpload(
    asUpload(makeTinyPdf([
      "Financial statements",
      "Operational update contained readable details for the quarter, including installations, trainings, customer activity, and management objectives."
    ]))
  );

  assert.deepEqual(result.diagnostics.financialImageHeavyPages, [1]);
  assert.deepEqual(result.diagnostics.pagesWithLittleText, [1]);
  assert.ok(result.diagnostics.pagesWithUsableText.includes(2));
});
