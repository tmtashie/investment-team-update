const path = require("path");

const MAX_PDF_UPLOAD_BYTES = 10 * 1024 * 1024;
const MAX_EXTRACTED_TEXT_LENGTH = 60000;
const VERY_SHORT_PAGE_TEXT_LENGTH = 80;

function looksFinancialPage(text) {
  return /\b(financial|balance sheet|income statement|statement of operations|cash flow|p&l|profit and loss|revenue|ebitda|runway|cash|debt|line of credit|loc|valuation|409a)\b/i.test(
    String(text || "")
  );
}

function buildDocumentDiagnostics({ pageCount, pages, combinedText }) {
  const pageMap = new Map((pages || []).map((page) => [Number(page.pageNumber), page]));
  const pagesWithUsableText = [];
  const pagesWithLittleText = [];
  const pagesSkipped = [];
  const financialImageHeavyPages = [];
  const totalPages = Number(pageCount) || pages.length || 0;

  for (let pageNumber = 1; pageNumber <= totalPages; pageNumber += 1) {
    const page = pageMap.get(pageNumber);
    const text = page ? asString(page.text, MAX_EXTRACTED_TEXT_LENGTH) : "";
    if (!text) {
      pagesSkipped.push(pageNumber);
      continue;
    }
    if (text.length < VERY_SHORT_PAGE_TEXT_LENGTH) {
      pagesWithLittleText.push(pageNumber);
    } else {
      pagesWithUsableText.push(pageNumber);
    }
    if (text.length < VERY_SHORT_PAGE_TEXT_LENGTH && looksFinancialPage(text)) {
      financialImageHeavyPages.push(pageNumber);
    }
  }

  return {
    totalPages,
    pagesWithUsableText,
    pagesWithLittleText,
    pagesAnalyzed: pages.map((page) => page.pageNumber).filter(Boolean),
    pagesSkipped,
    financialImageHeavyPages,
    extractedCharacterCount: String(combinedText || "").length
  };
}

function asString(value, maxLength = 2000) {
  return String(value || "").trim().slice(0, maxLength);
}

function isPdfFilename(filename) {
  return path.extname(asString(filename, 240)).toLowerCase() === ".pdf";
}

function normalizeBase64(value) {
  return asString(value, MAX_PDF_UPLOAD_BYTES * 2).replace(/^data:application\/pdf;base64,/i, "");
}

function validatePdfUpload({ filename, mimeType, fileData }) {
  const cleanFilename = asString(filename, 240);
  const cleanMimeType = asString(mimeType, 120).toLowerCase();
  const cleanFileData = normalizeBase64(fileData);

  if (!cleanFilename) {
    throw new Error("PDF filename is required.");
  }
  if (!isPdfFilename(cleanFilename)) {
    throw new Error("Only PDF files can be analyzed in this workflow.");
  }
  if (cleanMimeType && cleanMimeType !== "application/pdf" && cleanMimeType !== "application/octet-stream") {
    throw new Error("Only PDF files can be analyzed in this workflow.");
  }
  if (!cleanFileData) {
    throw new Error("PDF file data is required.");
  }

  const buffer = Buffer.from(cleanFileData, "base64");
  if (!buffer.length) {
    throw new Error("PDF file data is required.");
  }
  if (buffer.length > MAX_PDF_UPLOAD_BYTES) {
    throw new Error("PDF uploads are limited to 10 MB.");
  }
  if (buffer.subarray(0, 5).toString("latin1") !== "%PDF-") {
    throw new Error("The uploaded file is not a valid PDF.");
  }

  return {
    filename: cleanFilename,
    mimeType: cleanMimeType || "application/pdf",
    buffer
  };
}

function decodePdfLiteralString(value) {
  return String(value || "")
    .replace(/\\([nrtbf()\\])/g, (_match, escaped) => {
      const replacements = {
        n: "\n",
        r: "\r",
        t: "\t",
        b: "\b",
        f: "\f",
        "(": "(",
        ")": ")",
        "\\": "\\"
      };
      return replacements[escaped] || escaped;
    })
    .replace(/\\([0-7]{1,3})/g, (_match, octal) => String.fromCharCode(parseInt(octal, 8)));
}

function extractTextCommands(streamText) {
  const lines = [];
  const literalPattern = /\((?:\\.|[^\\)])*\)\s*Tj|\[(.*?)\]\s*TJ/gs;
  for (const match of streamText.matchAll(literalPattern)) {
    const raw = match[0];
    const strings = Array.from(raw.matchAll(/\((?:\\.|[^\\)])*\)/g))
      .map((item) => decodePdfLiteralString(item[0].slice(1, -1)))
      .filter(Boolean);
    if (strings.length) {
      lines.push(strings.join(""));
    }
  }
  return lines.join(" ").replace(/\s+/g, " ").trim();
}

function extractWithFallback(buffer, filename) {
  const raw = buffer.toString("latin1");
  const pages = [];
  const streamPattern = /stream\r?\n([\s\S]*?)\r?\nendstream/g;
  let pageNumber = 1;
  for (const match of raw.matchAll(streamPattern)) {
    const text = extractTextCommands(match[1]);
    if (text) {
      pages.push({ pageNumber, text });
    }
    pageNumber += 1;
  }
  const combinedText = pages.map((page) => `Page ${page.pageNumber}:\n${page.text}`).join("\n\n");
  if (!combinedText.trim()) {
    throw new Error("This PDF appears to be image-based or does not contain readable text.");
  }
  return {
    filename,
    pageCount: Math.max(pageNumber - 1, pages.length),
    pages,
    combinedText: combinedText.slice(0, MAX_EXTRACTED_TEXT_LENGTH),
    extractedTextLength: combinedText.length,
    diagnostics: buildDocumentDiagnostics({
      pageCount: Math.max(pageNumber - 1, pages.length),
      pages,
      combinedText
    })
  };
}

async function extractWithPdfParse(buffer, filename) {
  let pdfParse;
  try {
    pdfParse = require("pdf-parse");
  } catch (error) {
    return null;
  }

  const pages = [];
  const data = await pdfParse(buffer, {
    pagerender: async (pageData) => {
      const content = await pageData.getTextContent();
      const text = content.items
        .map((item) => asString(item.str, 2000))
        .filter(Boolean)
        .join(" ")
        .replace(/\s+/g, " ")
        .trim();
      pages.push({ pageNumber: pages.length + 1, text });
      return text;
    }
  });
  const readablePages = pages.filter((page) => page.text);
  const combinedText = readablePages.length
    ? readablePages.map((page) => `Page ${page.pageNumber}:\n${page.text}`).join("\n\n")
    : asString(data && data.text, MAX_EXTRACTED_TEXT_LENGTH);
  if (!combinedText.trim()) {
    throw new Error("This PDF appears to be image-based or does not contain readable text.");
  }
  return {
    filename,
    pageCount: Number(data && data.numpages) || readablePages.length || pages.length,
    pages: readablePages.length ? readablePages : [{ pageNumber: 1, text: combinedText }],
    combinedText: combinedText.slice(0, MAX_EXTRACTED_TEXT_LENGTH),
    extractedTextLength: combinedText.length,
    diagnostics: buildDocumentDiagnostics({
      pageCount: Number(data && data.numpages) || readablePages.length || pages.length,
      pages: readablePages.length ? readablePages : pages,
      combinedText
    })
  };
}

async function extractPdfTextFromUpload(upload) {
  const validated = validatePdfUpload(upload);
  try {
    const parsed = await extractWithPdfParse(validated.buffer, validated.filename);
    if (parsed) {
      return { ...parsed, buffer: validated.buffer };
    }
  } catch (error) {
    if (/password|encrypted/i.test(error.message || "")) {
      throw new Error("Encrypted or password-protected PDFs are not supported.");
    }
  }

  return { ...extractWithFallback(validated.buffer, validated.filename), buffer: validated.buffer };
}

module.exports = {
  MAX_PDF_UPLOAD_BYTES,
  validatePdfUpload,
  extractPdfTextFromUpload
};
