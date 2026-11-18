import * as pdfjsLib from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

export interface PdfTextItem {
  text: string;
  x: number;
  y: number;
  width: number;
  height: number;
  page: number;
}

export interface PdfLine {
  y: number;
  page: number;
  items: PdfTextItem[];
  text: string;
}

/**
 * Extracts text items and reconstructs visual lines sorted top-to-bottom, left-to-right.
 */
export async function extractPdfLines(file: File | ArrayBuffer): Promise<{
  lines: PdfLine[];
  rawText: string;
  pageCount: number;
}> {
  const buffer = file instanceof File ? await file.arrayBuffer() : file;
  const loadingTask = pdfjsLib.getDocument({ data: new Uint8Array(buffer) });
  const pdf = await loadingTask.promise;

  const allLines: PdfLine[] = [];

  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const textContent = await page.getTextContent();
    const items: PdfTextItem[] = [];

    for (const rawItem of textContent.items) {
      if ('str' in rawItem && rawItem.str.trim().length > 0) {
        const tx = rawItem.transform;
        items.push({
          text: rawItem.str,
          x: tx[4],
          y: tx[5],
          width: rawItem.width ?? 0,
          height: rawItem.height ?? 0,
          page: p,
        });
      }
    }

    // Sort top-to-bottom (higher Y first in PDF coordinates), then left-to-right
    items.sort((a, b) => {
      if (Math.abs(b.y - a.y) > 3.5) {
        return b.y - a.y;
      }
      return a.x - b.x;
    });

    // Cluster items into visual lines (tolerance 4pt)
    const pageLines: PdfLine[] = [];
    for (const item of items) {
      const existing = pageLines.find((line) => Math.abs(line.y - item.y) <= 4);
      if (existing) {
        existing.items.push(item);
      } else {
        pageLines.push({
          y: item.y,
          page: p,
          items: [item],
          text: '',
        });
      }
    }

    for (const line of pageLines) {
      line.items.sort((a, b) => a.x - b.x);
      line.text = line.items
        .map((i) => i.text.trim())
        .filter(Boolean)
        .join(' ');
      allLines.push(line);
    }
  }

  const rawText = allLines.map((l) => l.text).join('\n');
  return { lines: allLines, rawText, pageCount: pdf.numPages };
}

/**
 * Renders every page of a PDF to base64 JPEG images for Gemini multimodal extraction.
 */
export async function renderPdfPagesToImages(
  file: File,
  scale = 1.85,
  maxPages = 10
): Promise<Array<{ mimeType: string; data: string }>> {
  const buffer = await file.arrayBuffer();
  const loadingTask = pdfjsLib.getDocument({ data: new Uint8Array(buffer) });
  const pdf = await loadingTask.promise;

  const pagesToRender = Math.min(pdf.numPages, maxPages);
  const rendered: Array<{ mimeType: string; data: string }> = [];

  for (let p = 1; p <= pagesToRender; p++) {
    const page = await pdf.getPage(p);
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d');
    if (!context) continue;

    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);

    await page.render({
      canvasContext: context,
      viewport,
      canvas,
    }).promise;

    const dataUrl = canvas.toDataURL('image/jpeg', 0.86);
    const base64 = dataUrl.split(',')[1] ?? '';
    rendered.push({
      mimeType: 'image/jpeg',
      data: base64,
    });
  }

  return rendered;
}
