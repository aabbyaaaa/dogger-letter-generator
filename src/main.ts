import "./styles.css";
import { toPng } from "html-to-image";
import jsPDF from "jspdf";
import {
  AlignmentType,
  Document,
  HorizontalPositionRelativeFrom,
  ImageRun,
  Packer,
  PageOrientation,
  Paragraph,
  SectionType,
  TextRun,
  VerticalPositionRelativeFrom,
} from "docx";

import stationery01 from "../2026-德記-英文信籤-01.jpg";
import stationery02 from "../2026-德記-英文信籤-02.jpg";
import signatureImage from "../TI簽名檔.jpg";

type TemplateId = "01" | "02";

type LetterState = {
  templateId: TemplateId;
  date: string;
  content: string;
  includeSignature: boolean;
  signatureOffsetX: number;
  signatureOffsetY: number;
  signatureWidth: number;
  signerName: string;
  signerTitle: string;
};

type TemplateConfig = {
  id: TemplateId;
  label: string;
  image: string;
  contentLeft: number;
  contentTop: number;
  contentWidth: number;
  signatureTop: number;
  signatureLeft: number;
  dateTop: number;
  dateRight: number;
};

const PAGE_WIDTH = 794;
const PAGE_HEIGHT = 1123;
const LETTER_FONT_SIZE = 16;
const LETTER_LINE_HEIGHT = 28;
const SIGNATURE_ASPECT_RATIO = 363 / 543;
const SIGNER_NAME_LINE_HEIGHT = 23;
const SIGNATURE_TITLE_GAP = 6;
const DATE_LINE_HEIGHT = 20;
const A4_WIDTH_MM = 210;
const A4_HEIGHT_MM = 297;
const A4_WIDTH_TWIP = 11906;
const A4_HEIGHT_TWIP = 16838;
const TWIP_PER_PX = 15;

const templates: Record<TemplateId, TemplateConfig> = {
  "01": {
    id: "01",
    label: "英文信籤 01",
    image: stationery01,
    contentLeft: 84,
    contentTop: 152,
    contentWidth: 660,
    signatureTop: 880,
    signatureLeft: 84,
    dateTop: 988,
    dateRight: 94,
  },
  "02": {
    id: "02",
    label: "英文信籤 02",
    image: stationery02,
    contentLeft: 92,
    contentTop: 164,
    contentWidth: 650,
    signatureTop: 880,
    signatureLeft: 88,
    dateTop: 988,
    dateRight: 94,
  },
};

const sampleContent = `Subject : Notification of Company Name and Office Address Change

Dear Valued Partner,

We would like to inform you that effective from May 11, 2026, our company name has officially changed from:
DOGGER INSTRUMENTS CO., LTD.
to
DOGGER SCIENTIFIC CO., LTD.
At the same time, our office address has been relocated to:
DOGGER SCIENTIFIC CO., LTD.
5F, No. 310, Jianguo 1st Rd. Xinzhuang Dist., New Taipei City 24207, Taiwan
Please note that our shipping address remains unchanged as below:
No. 12, Lane 172, Jiun-Ying St. Shu-Lin Dist., New Taipei City 23863,Taiwan

Except for the above changes, all other company information, including our telephone numbers, email addresses, and shipping address, will remain unchanged.
We kindly ask you to update your records accordingly.
Thank you very much for your continued support and cooperation. Should you have any questions, please feel free to contact us.

Best regards,
DOGGER SCIENTIFIC CO., LTD.`;

const state: LetterState = {
  templateId: "01",
  date: toDateInputValue(new Date()),
  content: sampleContent,
  includeSignature: true,
  signatureOffsetX: 88,
  signatureOffsetY: 0,
  signatureWidth: 170,
  signerName: "Ti Yen",
  signerTitle: "Manager / International business",
};

const app = document.querySelector<HTMLDivElement>("#app");

if (!app) {
  throw new Error("App root was not found.");
}

app.innerHTML = `
  <section class="workspace">
    <aside class="control-panel" aria-label="信件設定">
      <div class="brand-lockup">
        <span class="brand-mark">D</span>
        <div>
          <p class="eyebrow">DOGGER INTERNAL TOOL</p>
          <h1>德記信件產生器</h1>
        </div>
      </div>

      <div class="field-group">
        <label class="field-label">信紙版本</label>
        <div class="segmented" role="radiogroup" aria-label="信紙版本">
          <label>
            <input type="radio" name="template" value="01" checked />
            <span>01</span>
          </label>
          <label>
            <input type="radio" name="template" value="02" />
            <span>02</span>
          </label>
        </div>
      </div>

      <label class="field-group">
        <span class="field-label">日期</span>
        <input id="dateInput" class="control-input" type="date" value="${state.date}" />
      </label>

      <label class="field-group field-grow">
        <span class="field-label">信件內容</span>
        <textarea id="contentInput" class="content-input" spellcheck="false">${escapeHtml(state.content)}</textarea>
      </label>

      <label class="signature-toggle">
        <input id="includeSignatureInput" type="checkbox" checked />
        <span class="toggle-box" aria-hidden="true"></span>
        <span class="toggle-copy">
          <strong>放入手寫簽名</strong>
          <span>預設使用 TI 簽名檔；取消勾選時只保留姓名與職稱。</span>
        </span>
      </label>

      <div class="signature-adjustments" aria-label="手寫簽名位置微調">
        <label class="field-group">
          <span class="field-label">簽名左右</span>
          <input id="signatureXInput" class="control-input" type="number" min="-120" max="360" step="1" value="${state.signatureOffsetX}" />
        </label>
        <label class="field-group">
          <span class="field-label">上下微調</span>
          <input id="signatureYInput" class="control-input" type="number" min="-80" max="80" step="1" value="${state.signatureOffsetY}" />
        </label>
        <label class="field-group">
          <span class="field-label">簽名大小</span>
          <input id="signatureWidthInput" class="control-input" type="number" min="80" max="320" step="1" value="${state.signatureWidth}" />
        </label>
      </div>

      <div class="signature-fields">
        <label class="field-group">
          <span class="field-label">姓名</span>
          <input id="signerNameInput" class="control-input" type="text" value="${escapeHtml(state.signerName)}" />
        </label>
        <label class="field-group">
          <span class="field-label">職稱</span>
          <input id="signerTitleInput" class="control-input" type="text" value="${escapeHtml(state.signerTitle)}" />
        </label>
      </div>

      <div class="actions">
        <button id="pdfButton" class="primary-action" type="button">
          <span aria-hidden="true">PDF</span>
          下載 PDF
        </button>
        <button id="wordButton" class="secondary-action" type="button">
          <span aria-hidden="true">DOCX</span>
          下載 Word
        </button>
      </div>

      <p id="statusText" class="status-text" role="status">已準備好產生文件。</p>
    </aside>

    <section class="preview-shell" aria-label="信件預覽">
      <header class="preview-header">
        <div>
          <p class="eyebrow">LIVE A4 PREVIEW</p>
          <h2>預覽</h2>
        </div>
        <div id="pageCount" class="page-count">1 page</div>
      </header>
      <div id="previewPages" class="preview-pages"></div>
    </section>
  </section>
`;

const previewPages = mustGet<HTMLDivElement>("previewPages");
const pageCount = mustGet<HTMLDivElement>("pageCount");
const statusText = mustGet<HTMLParagraphElement>("statusText");
const dateInput = mustGet<HTMLInputElement>("dateInput");
const contentInput = mustGet<HTMLTextAreaElement>("contentInput");
const includeSignatureInput = mustGet<HTMLInputElement>("includeSignatureInput");
const signatureXInput = mustGet<HTMLInputElement>("signatureXInput");
const signatureYInput = mustGet<HTMLInputElement>("signatureYInput");
const signatureWidthInput = mustGet<HTMLInputElement>("signatureWidthInput");
const signerNameInput = mustGet<HTMLInputElement>("signerNameInput");
const signerTitleInput = mustGet<HTMLInputElement>("signerTitleInput");
const pdfButton = mustGet<HTMLButtonElement>("pdfButton");
const wordButton = mustGet<HTMLButtonElement>("wordButton");

document.querySelectorAll<HTMLInputElement>('input[name="template"]').forEach((input) => {
  input.addEventListener("change", () => {
    if (input.checked) {
      state.templateId = input.value as TemplateId;
      renderPreview();
    }
  });
});

dateInput.addEventListener("input", () => {
  state.date = dateInput.value;
  renderPreview();
});

contentInput.addEventListener("input", () => {
  state.content = contentInput.value;
  renderPreview();
});

includeSignatureInput.addEventListener("change", () => {
  state.includeSignature = includeSignatureInput.checked;
  renderPreview();
});

signatureXInput.addEventListener("input", () => {
  state.signatureOffsetX = numberValue(signatureXInput.value, state.signatureOffsetX);
  renderPreview();
});

signatureYInput.addEventListener("input", () => {
  state.signatureOffsetY = numberValue(signatureYInput.value, state.signatureOffsetY);
  renderPreview();
});

signatureWidthInput.addEventListener("input", () => {
  state.signatureWidth = numberValue(signatureWidthInput.value, state.signatureWidth);
  renderPreview();
});

signerNameInput.addEventListener("input", () => {
  state.signerName = signerNameInput.value;
  renderPreview();
});

signerTitleInput.addEventListener("input", () => {
  state.signerTitle = signerTitleInput.value;
  renderPreview();
});

pdfButton.addEventListener("click", () => {
  void downloadPdf();
});

wordButton.addEventListener("click", () => {
  void downloadWord();
});

renderPreview();

function renderPreview(): void {
  const template = templates[state.templateId];
  const pages = paginateContent(state.content, template);

  previewPages.innerHTML = pages
    .map((lines, index) => renderPage(template, lines, index, pages.length))
    .join("");

  const total = pages.length;
  pageCount.textContent = `${total} page${total > 1 ? "s" : ""}`;
}

function renderPage(
  template: TemplateConfig,
  lines: string[],
  index: number,
  total: number,
): string {
  const isLastPage = index === total - 1;
  const textLines = lines
    .map((line) => `<div class="letter-line">${line ? escapeHtml(line) : "&nbsp;"}</div>`)
    .join("");

  return `
    <article
      class="letter-page"
      data-page="${index + 1}"
      style="background-image: url('${template.image}')"
      aria-label="第 ${index + 1} 頁"
    >
      <div
        class="letter-content"
        style="
          left: ${template.contentLeft}px;
          top: ${template.contentTop}px;
          width: ${template.contentWidth}px;
        "
      >${textLines}</div>

      ${
        isLastPage
          ? `
            <section
              class="letter-signature ${state.includeSignature ? "" : "without-handwriting"}"
              style="left: ${template.signatureLeft}px; top: ${signatureBlockTop(template)}px"
              aria-label="簽名"
            >
              <div class="signature-name-row">
                <strong>${escapeHtml(state.signerName || " ")}</strong>
                ${
                  state.includeSignature
                    ? `
                      <img
                        src="${signatureImage}"
                        alt="Ti 手寫簽名"
                        style="
                          margin-left: ${state.signatureOffsetX}px;
                          width: ${state.signatureWidth}px;
                        "
                      />
                    `
                    : ""
                }
              </div>
              <span>${escapeHtml(state.signerTitle || " ")}</span>
            </section>
            <p
              class="letter-date"
              style="right: ${template.dateRight}px; top: ${template.dateTop}px"
            ><span>Date :</span> ${escapeHtml(displayDate(state.date))}</p>
          `
          : ""
      }
    </article>
  `;
}

function paginateContent(content: string, template: TemplateConfig): string[][] {
  const availableHeight = signatureVisualTop(template) - template.contentTop - 18;
  const linesPerPage = Math.max(1, Math.floor(availableHeight / LETTER_LINE_HEIGHT));
  const wrappedLines = wrapText(content.trimEnd(), template.contentWidth);

  if (wrappedLines.length === 0) {
    return [[]];
  }

  const pages: string[][] = [];
  for (let index = 0; index < wrappedLines.length; index += linesPerPage) {
    pages.push(wrappedLines.slice(index, index + linesPerPage));
  }

  return pages;
}

function wrapText(text: string, maxWidth: number): string[] {
  const context = getMeasureContext();
  context.font = `${LETTER_FONT_SIZE}px "Noto Sans TC", "Microsoft JhengHei", "Segoe UI", sans-serif`;

  const paragraphs = text.split(/\r?\n/);
  const lines: string[] = [];

  paragraphs.forEach((paragraph) => {
    if (!paragraph) {
      lines.push("");
      return;
    }

    let current = "";

    splitWrapTokens(paragraph).forEach((token) => {
      const next = current + token;
      if (!current || context.measureText(next).width <= maxWidth) {
        current = next;
        return;
      }

      lines.push(current.trimEnd());
      current = token.trimStart();

      if (context.measureText(current).width > maxWidth) {
        const broken = breakLongToken(current, maxWidth, context);
        lines.push(...broken.complete);
        current = broken.remainder;
      }
    });

    if (current) {
      lines.push(current.trimEnd());
    }
  });

  return lines;
}

function splitWrapTokens(paragraph: string): string[] {
  return paragraph.match(/\s+|[^\s]+/g) ?? [];
}

function breakLongToken(
  token: string,
  maxWidth: number,
  context: CanvasRenderingContext2D,
): { complete: string[]; remainder: string } {
  const complete: string[] = [];
  let current = "";

  Array.from(token).forEach((char) => {
    const next = current + char;
    if (!current || context.measureText(next).width <= maxWidth) {
      current = next;
      return;
    }
    complete.push(current);
    current = char;
  });

  return { complete, remainder: current };
}

function getMeasureContext(): CanvasRenderingContext2D {
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");
  if (!context) {
    throw new Error("Canvas is not supported in this browser.");
  }
  return context;
}

async function downloadPdf(): Promise<void> {
  setBusy(true, "正在產生 PDF...");
  try {
    await document.fonts.ready;
    const pages = Array.from(document.querySelectorAll<HTMLElement>(".letter-page"));
    const pdf = new jsPDF({
      orientation: "portrait",
      unit: "mm",
      format: "a4",
      compress: true,
    });

    for (const [index, page] of pages.entries()) {
      const png = await toPng(page, {
        cacheBust: true,
        pixelRatio: 2,
        width: PAGE_WIDTH,
        height: PAGE_HEIGHT,
        style: {
          width: `${PAGE_WIDTH}px`,
          height: `${PAGE_HEIGHT}px`,
          transform: "none",
        },
      });

      if (index > 0) {
        pdf.addPage("a4", "portrait");
      }
      pdf.addImage(png, "PNG", 0, 0, A4_WIDTH_MM, A4_HEIGHT_MM);
    }

    pdf.save(`德記信件-${fileDate()}.pdf`);
    setBusy(false, "PDF 已產生。");
  } catch (error) {
    console.error(error);
    setBusy(false, "PDF 產生失敗，請稍後再試。");
  }
}

async function downloadWord(): Promise<void> {
  setBusy(true, "正在產生 Word...");
  try {
    const template = templates[state.templateId];
    const pages = paginateContent(state.content, template);
    const [paperData, signatureData] = await Promise.all([
      fetchArrayBuffer(template.image),
      fetchArrayBuffer(signatureImage),
    ]);

    const documentSections = pages.map((lines, index) => ({
      properties: {
        type: index === 0 ? SectionType.CONTINUOUS : SectionType.NEXT_PAGE,
        page: {
          size: {
            width: A4_WIDTH_TWIP,
            height: A4_HEIGHT_TWIP,
            orientation: PageOrientation.PORTRAIT,
          },
          margin: {
            top: Math.round(template.contentTop * TWIP_PER_PX),
            right: Math.round((PAGE_WIDTH - template.contentLeft - template.contentWidth) * TWIP_PER_PX),
            bottom: 900,
            left: Math.round(template.contentLeft * TWIP_PER_PX),
          },
        },
      },
      children: [
        backgroundImageRun(paperData),
        ...lines.map((line) => textParagraph(line)),
        ...(index === pages.length - 1
          ? signatureWordBlock(signatureData, lines.length, template)
          : []),
      ],
    }));

    const doc = new Document({
      creator: "DOGGER Scientific Co., Ltd.",
      title: "Dogger Letter",
      styles: {
        paragraphStyles: [
          {
            id: "LetterBody",
            name: "Letter Body",
            basedOn: "Normal",
            next: "LetterBody",
            run: {
              size: 25,
              font: "Microsoft JhengHei",
              color: "231f20",
            },
            paragraph: {
              spacing: { before: 0, after: 120, line: 390 },
            },
          },
        ],
      },
      sections: documentSections,
    });

    const blob = await Packer.toBlob(doc);
    saveBlob(blob, `德記信件-${fileDate()}.docx`);
    setBusy(false, "Word 已產生。");
  } catch (error) {
    console.error(error);
    setBusy(false, "Word 產生失敗，請稍後再試。");
  }
}

function backgroundImageRun(data: ArrayBuffer): Paragraph {
  return new Paragraph({
    spacing: { before: 0, after: 0 },
    children: [
      new ImageRun({
        type: "jpg",
        data,
        transformation: {
          width: PAGE_WIDTH,
          height: PAGE_HEIGHT,
        },
        floating: {
          behindDocument: true,
          allowOverlap: true,
          horizontalPosition: {
            relative: HorizontalPositionRelativeFrom.PAGE,
            offset: 0,
          },
          verticalPosition: {
            relative: VerticalPositionRelativeFrom.PAGE,
            offset: 0,
          },
        },
      }),
    ],
  });
}

function textParagraph(text: string): Paragraph {
  return new Paragraph({
    style: "LetterBody",
    children: [
      new TextRun({
        text: text || " ",
        font: {
          ascii: "Calibri",
          hAnsi: "Calibri",
          eastAsia: "Microsoft JhengHei",
          cs: "Microsoft JhengHei",
        },
      }),
    ],
  });
}

function signatureWordBlock(
  signatureData: ArrayBuffer,
  contentLineCount: number,
  template: TemplateConfig,
): Paragraph[] {
  const linesBeforeSignature = Math.max(
    1,
    Math.floor((signatureBlockTop(template) - template.contentTop) / LETTER_LINE_HEIGHT) -
      contentLineCount -
      1,
  );

  const spacers = Array.from({ length: linesBeforeSignature }, () =>
    new Paragraph({ style: "LetterBody", children: [new TextRun(" ")] }),
  );

  return [
    ...spacers,
    new Paragraph({
      style: "LetterBody",
      children: [
        new TextRun({
          text: state.signerName || " ",
          bold: true,
          size: 25,
        }),
        ...(state.includeSignature
          ? [
              new TextRun({
                text: "    ",
              }),
              new ImageRun({
                type: "jpg",
                data: signatureData,
                transformation: {
                  width: Math.round(state.signatureWidth * 0.9),
                  height: Math.round(state.signatureWidth * 0.6),
                },
              }),
            ]
          : []),
      ],
    }),
    new Paragraph({
      style: "LetterBody",
      children: [new TextRun(state.signerTitle || " ")],
    }),
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      spacing: { before: 120, after: 0 },
      children: [
        new TextRun({
          text: "Date :",
          color: "00577d",
          size: 25,
        }),
        new TextRun({
          text: ` ${displayDate(state.date)}`,
          color: "231f20",
          size: 25,
        }),
      ],
    }),
  ];
}

async function fetchArrayBuffer(url: string): Promise<ArrayBuffer> {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Unable to load asset: ${url}`);
  }
  return response.arrayBuffer();
}

function saveBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function setBusy(isBusy: boolean, message: string): void {
  pdfButton.disabled = isBusy;
  wordButton.disabled = isBusy;
  statusText.textContent = message;
}

function displayDate(value: string): string {
  if (!value) {
    return "";
  }
  return value.replaceAll("-", "/");
}

function fileDate(): string {
  const date = displayDate(state.date) || displayDate(toDateInputValue(new Date()));
  return date.replaceAll("/", "");
}

function toDateInputValue(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function mustGet<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) {
    throw new Error(`Element #${id} was not found.`);
  }
  return element as T;
}

function numberValue(value: string, fallback: number): number {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function signatureImageHeight(): number {
  return Math.round(state.signatureWidth * SIGNATURE_ASPECT_RATIO);
}

function signatureBlockTop(template: TemplateConfig): number {
  return Math.round(
    template.dateTop + DATE_LINE_HEIGHT - signatureBlockHeight() + state.signatureOffsetY,
  );
}

function signatureVisualTop(template: TemplateConfig): number {
  return signatureBlockTop(template);
}

function signatureBlockHeight(): number {
  const nameRowHeight = state.includeSignature
    ? Math.max(SIGNER_NAME_LINE_HEIGHT, signatureImageHeight())
    : SIGNER_NAME_LINE_HEIGHT;
  return Math.round(nameRowHeight + SIGNATURE_TITLE_GAP + SIGNER_NAME_LINE_HEIGHT);
}

function escapeHtml(value: string): string {
  const div = document.createElement("div");
  div.textContent = value;
  return div.innerHTML;
}
