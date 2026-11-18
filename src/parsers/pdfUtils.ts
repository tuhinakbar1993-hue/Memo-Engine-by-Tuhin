import * as pdfjsLib from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

// Pre-warm a single PDFWorker instance so first upload has zero worker startup lag
let sharedWorker: pdfjsLib.PDFWorker | null = null;
function getSharedPdfWorker(): pdfjsLib.PDFWorker {
  if (!sharedWorker) {
    sharedWorker = new pdfjsLib.PDFWorker();
  }
  return sharedWorker;
}

// Pre-warm immediately in browser idle time
if (typeof window !== 'undefined') {
  const warm = () => {
    try {
      getSharedPdfWorker();
    } catch {
      // Ignore pre-warm errors
    }
  };
  if ('requestIdleCallback' in window) {
    window.requestIdleCallback(warm);
  } else {
    setTimeout(warm, 150);
  }
}

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
 * Extracts text items across all pages in parallel and reconstructs visual lines
 * sorted top-to-bottom, left-to-right.
 */
export async function extractPdfLines(file: File | ArrayBuffer): Promise<{
  lines: PdfLine[];
  rawText: string;
  pageCount: number;
}> {
  const buffer = file instanceof File ? await file.arrayBuffer() : file;
  const loadingTask = pdfjsLib.getDocument({
    data: new Uint8Array(buffer),
    worker: getSharedPdfWorker(),
  });
  const pdf = await loadingTask.promise;

  const pagePromises = Array.from({ length: pdf.numPages }, async (_, idx) => {
    const p = idx + 1;
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

    items.sort((a, b) => {
      if (Math.abs(b.y - a.y) > 3.5) {
        return b.y - a.y;
      }
      return a.x - b.x;
    });

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
    }

    return pageLines;
  });

  const perPageLines = await Promise.all(pagePromises);
  const allLines = perPageLines.flat();
  const rawText = allLines.map((l) => l.text).join('\n');

  return { lines: allLines, rawText, pageCount: pdf.numPages };
}

/**
 * Renders PDF pages in parallel to optimally sized JPEG images (max 1400px width)
 * for ultra-fast network transfer and Gemini multimodal extraction.
 */
export async function renderPdfPagesToImages(
  file: File,
  maxWidth = 1400,
  maxPages = 10
): Promise<Array<{ mimeType: string; data: string }>> {
  const buffer = await file.arrayBuffer();
  const loadingTask = pdfjsLib.getDocument({
    data: new Uint8Array(buffer),
    worker: getSharedPdfWorker(),
  });
  const pdf = await loadingTask.promise;

  const pagesToRender = Math.min(pdf.numPages, maxPages);

  const renderPromises = Array.from({ length: pagesToRender }, async (_, idx) => {
    const p = idx + 1;
    const page = await pdf.getPage(p);
    const baseViewport = page.getViewport({ scale: 1 });
    const targetScale = Math.min(1.65, Math.max(1.15, maxWidth / baseViewport.width));
    const viewport = page.getViewport({ scale: targetScale });

    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d');
    if (!context) return null;

    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);

    await page.render({
      canvasContext: context,
      viewport,
      canvas,
    }).promise;

    const dataUrl = canvas.toDataURL('image/jpeg', 0.78);
    const base64 = dataUrl.split(',')[1] ?? '';
    return {
      mimeType: 'image/jpeg',
      data: base64,
    };
  });

  const results = await Promise.all(renderPromises);
  return results.filter((r): r is { mimeType: string; data: string } => r !== null);
}
