import * as XLSX from 'xlsx';
import type {
  AccountStatementData,
  BranchProposalData,
  CibReportData,
  CostOfFundData,
  DocumentKey,
  OutstandingData,
  OverdueData,
  PartyPerformanceData,
} from '../types/memo';
import { cleanText, parseIndianNumber } from '../utils/numberFormat';
import { extractPdfLines, PdfTextItem, renderPdfPagesToImages } from './pdfUtils';

const parseCache = new Map<string, unknown>();

function getFileSignature(prefix: string, file: File): string {
  return `${prefix}:${file.name}:${file.size}:${file.lastModified}`;
}

export function getCachedParse<T>(prefix: string, file: File): T | undefined {
  return parseCache.get(getFileSignature(prefix, file)) as T | undefined;
}

export function setCachedParse<T>(prefix: string, file: File, value: T): T {
  parseCache.set(getFileSignature(prefix, file), value);
  return value;
}

function findNearestItemByX(items: PdfTextItem[], targetX: number, maxDist = 65): PdfTextItem | undefined {
  let best: PdfTextItem | undefined;
  let bestDist = Infinity;
  for (const item of items) {
    const center = item.x + item.width / 2;
    const dist = Math.abs(center - targetX);
    if (dist < bestDist && dist <= maxDist) {
      bestDist = dist;
      best = item;
    }
  }
  return best;
}

export type StageCallback = (stage: string) => void;

/**
 * 1. Branch Proposal (Scanned PDF or Images -> Gemini multimodal API route)
 */
export async function parseBranchProposalFile(
  file: File,
  onStage?: StageCallback
): Promise<BranchProposalData> {
  onStage?.('Checking cache...');
  const cached = getCachedParse<BranchProposalData>('bp', file);
  if (cached) {
    onStage?.('Loaded from cache');
    return cached;
  }

  let pages: Array<{ mimeType: string; data: string }> = [];

  if (file.type.startsWith('image/')) {
    onStage?.('Reading image...');
    const dataUrl = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result ?? ''));
      reader.onerror = () => reject(new Error('Failed to read image file.'));
      reader.readAsDataURL(file);
    });
    pages = [{ mimeType: file.type, data: dataUrl.split(',')[1] ?? '' }];
  } else {
    onStage?.('Rendering scanned pages...');
    pages = await renderPdfPagesToImages(file, 1400, 8);
  }

  if (pages.length === 0) {
    throw new Error('Could not render any pages from the Branch Proposal file.');
  }

  onStage?.(`Analyzing ${pages.length} page${pages.length > 1 ? 's' : ''} with Gemini...`);
  const response = await fetch('/api/gemini/extract-branch-proposal', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ pages }),
  });

  if (!response.ok) {
    const errBody = await response.json().catch(() => ({}));
    throw new Error(errBody.error || `Branch Proposal extraction failed (${response.status}).`);
  }

  onStage?.('Finalizing extracted fields...');
  const json = await response.json();
  return setCachedParse('bp', file, json.data as BranchProposalData);
}

/**
 * 2. Outstanding PDF (Grand Total row: Princ Out, Charge, Profit Out, Compensation, Unapplied Pro, NET Out)
 */
export async function parseOutstandingPdf(
  file: File,
  onStage?: StageCallback
): Promise<{
  data: OutstandingData;
  rawText: string;
}> {
  const cached = getCachedParse<{ data: OutstandingData; rawText: string }>('outstanding', file);
  if (cached) {
    onStage?.('Loaded from cache');
    return cached;
  }

  onStage?.('Reading PDF text...');
  const { lines, rawText } = await extractPdfLines(file);
  onStage?.('Extracting balances...');

  // Locate header line containing Princ Out and NET Out
  const headerLine = lines.find(
    (l) => /princ\s*out/i.test(l.text) && /net\s*out/i.test(l.text)
  );
  const grandTotalLine = lines.find((l) => /grand\s*total/i.test(l.text));

  let princOutBDT: number | null = null;
  let chargeBDT: number | null = null;
  let profitOutBDT: number | null = null;
  let compensationBDT: number | null = null;
  let unappliedProBDT: number | null = null;
  let netOutBDT: number | null = null;

  if (headerLine && grandTotalLine) {
    const findColX = (pattern: RegExp): number | null => {
      const match = headerLine.items.find((it) => pattern.test(it.text));
      return match ? match.x + match.width / 2 : null;
    };

    const princX = findColX(/princ\s*out/i);
    const profitX = findColX(/profit\s*out/i);
    const compX = findColX(/compensation/i);
    const unappX = findColX(/unapplied/i);
    const chargeX = findColX(/charge/i);
    const netX = findColX(/net\s*out/i);

    const numericItems = grandTotalLine.items.filter(
      (it) => parseIndianNumber(it.text) !== null
    );

    if (princX !== null) {
      const m = findNearestItemByX(numericItems, princX);
      if (m) princOutBDT = parseIndianNumber(m.text);
    }
    if (profitX !== null) {
      const m = findNearestItemByX(numericItems, profitX);
      if (m) profitOutBDT = parseIndianNumber(m.text);
    }
    if (compX !== null) {
      const m = findNearestItemByX(numericItems, compX);
      if (m) compensationBDT = parseIndianNumber(m.text);
    }
    if (unappX !== null) {
      const m = findNearestItemByX(numericItems, unappX);
      if (m) unappliedProBDT = parseIndianNumber(m.text);
    }
    if (chargeX !== null) {
      const m = findNearestItemByX(numericItems, chargeX);
      if (m) chargeBDT = parseIndianNumber(m.text);
    }
    if (netX !== null) {
      const m = findNearestItemByX(numericItems, netX);
      if (m) netOutBDT = parseIndianNumber(m.text);
    }
  }

  // Positional fallback if header items were merged into single strings
  if (grandTotalLine && netOutBDT === null) {
    const nums = grandTotalLine.text
      .replace(/grand\s*total\s*:?/i, '')
      .trim()
      .split(/\s+/)
      .map((tok) => parseIndianNumber(tok))
      .filter((n): n is number => n !== null);
    // Expected column order: Disburse, Princ Out, Profit Out, Compensation, Unapplied Pro, Charge, Rebate, NET Out
    if (nums.length >= 8) {
      princOutBDT = princOutBDT ?? nums[1];
      profitOutBDT = profitOutBDT ?? nums[2];
      compensationBDT = compensationBDT ?? nums[3];
      unappliedProBDT = unappliedProBDT ?? nums[4];
      chargeBDT = chargeBDT ?? nums[5];
      netOutBDT = netOutBDT ?? nums[7];
    }
  }

  return setCachedParse('outstanding', file, {
    data: {
      princOutBDT,
      chargeBDT,
      profitOutBDT,
      compensationBDT,
      unappliedProBDT,
      netOutBDT,
    },
    rawText,
  });
}

/**
 * 3. Party Performance PDF (Product Name -> #5#/#55#, Customer Total Recovery Amt -> #60#, Print Date -> #68#)
 */
export async function parsePartyPerformancePdf(
  file: File,
  onStage?: StageCallback
): Promise<{
  data: PartyPerformanceData;
  rawText: string;
}> {
  const cached = getCachedParse<{ data: PartyPerformanceData; rawText: string }>('partyPerf', file);
  if (cached) {
    onStage?.('Loaded from cache');
    return cached;
  }

  onStage?.('Reading PDF text...');
  const { lines, rawText } = await extractPdfLines(file);
  onStage?.('Extracting performance figures...');

  let productName = '';
  let printDate = '';
  let recoveryAmtBDT: number | null = null;

  // Print Date
  for (const line of lines) {
    const m = line.text.match(/Print\s*Date\s*:\s*([0-9./-]+)/i);
    if (m) {
      printDate = cleanText(m[1]);
      break;
    }
  }

  // Product Name
  for (const line of lines) {
    const m = line.text.match(/Product\s*Name\s*:?\s*(.+)/i);
    if (m) {
      productName = cleanText(m[1]);
      break;
    }
  }

  // Recovery Amt. column X-coordinate
  const recHeaderLine = lines.find((l) => /recovery\s*amt/i.test(l.text));
  let recoveryColX: number | null = null;
  if (recHeaderLine) {
    const recItem = recHeaderLine.items.find((it) => /recovery/i.test(it.text));
    if (recItem) {
      recoveryColX = recItem.x + recItem.width / 2;
    }
  }

  // Find Customer Total section: look for the line with "Customer Total" and the 5 stacked lines around it
  const custTotalIdx = lines.findIndex((l) => /customer\s*total/i.test(l.text));
  if (custTotalIdx !== -1) {
    // Collect lines in a tight window around Customer Total (2 before to 3 after), skipping Product Total & footer
    const windowLines = lines
      .slice(Math.max(0, custTotalIdx - 2), custTotalIdx + 4)
      .filter((l) => !/product\s*total|prepared\s*by/i.test(l.text));
    const recoveryNumbers: Array<{ y: number; val: number }> = [];

    for (const wl of windowLines) {
      if (recoveryColX !== null) {
        const match = findNearestItemByX(
          wl.items.filter((it) => parseIndianNumber(it.text) !== null),
          recoveryColX,
          55
        );
        if (match) {
          const v = parseIndianNumber(match.text);
          if (v !== null) recoveryNumbers.push({ y: wl.y, val: v });
        }
      } else {
        const nums = wl.items
          .map((it) => parseIndianNumber(it.text))
          .filter((n): n is number => n !== null);
        if (nums.length >= 2) {
          recoveryNumbers.push({ y: wl.y, val: nums[1] });
        }
      }
    }

    // The 5th/lowest line in Customer Total block is "Total" of Principal/Profit/Others/Compensation/Total
    if (recoveryNumbers.length > 0) {
      recoveryNumbers.sort((a, b) => a.y - b.y); // lowest Y is the bottom Total row
      recoveryAmtBDT = recoveryNumbers[0].val;
    }
  }

  return setCachedParse('partyPerf', file, {
    data: {
      productName,
      recoveryAmtBDT,
      printDate,
    },
    rawText,
  });
}

/**
 * 4. Overdue PDF (Grand Total of Overdue Amount -> #59#)
 */
export async function parseOverduePdf(
  file: File,
  onStage?: StageCallback
): Promise<{
  data: OverdueData;
  rawText: string;
}> {
  const cached = getCachedParse<{ data: OverdueData; rawText: string }>('overdue', file);
  if (cached) {
    onStage?.('Loaded from cache');
    return cached;
  }

  onStage?.('Reading PDF text...');
  const { lines, rawText } = await extractPdfLines(file);
  onStage?.('Extracting overdue total...');

  let overdueAmountBDT: number | null = null;
  const grandTotalLine = lines.find((l) => /grand\s*total/i.test(l.text));

  if (grandTotalLine) {
    const nums = grandTotalLine.items
      .map((it) => parseIndianNumber(it.text))
      .filter((n): n is number => n !== null);
    if (nums.length > 0) {
      overdueAmountBDT = nums[nums.length - 1];
    }
  }

  return setCachedParse('overdue', file, {
    data: { overdueAmountBDT },
    rawText,
  });
}

/**
 * 5. Account Statement PDF (#2# Account number, #36# A/C No., #3# Branch name 2nd line, #69# Customer ID, #35# Last row Balance)
 */
export async function parseAccountStatementPdf(
  file: File,
  onStage?: StageCallback
): Promise<{
  data: AccountStatementData;
  rawText: string;
}> {
  const cached = getCachedParse<{ data: AccountStatementData; rawText: string }>('acctStmt', file);
  if (cached) {
    onStage?.('Loaded from cache');
    return cached;
  }

  onStage?.('Reading PDF text...');
  const { lines, rawText } = await extractPdfLines(file);
  onStage?.('Extracting account & balance...');

  // Branch name: 2nd line of the document (excluding standalone Print Date tokens on the right margin)
  let branchName = '';
  const page1Lines = lines.filter((l) => l.page === 1);
  if (page1Lines.length >= 2) {
    // Filter out any Print Date token on the header lines
    const secondLineCleaned = page1Lines[1].items
      .filter((it) => !/print\s*date/i.test(it.text))
      .map((it) => it.text.trim())
      .join(' ');
    branchName = cleanText(secondLineCleaned);
  }

  let customerId = '';
  let acNo = '';

  for (const line of lines) {
    const cidMatch = line.text.match(/Customer\s*ID\s*:\s*([A-Za-z0-9-]+)/i);
    if (cidMatch && !customerId) {
      customerId = cleanText(cidMatch[1]);
    }
    const acMatch = line.text.match(/A\/C\s*No\.?\s*:\s*([A-Za-z0-9-]+)/i);
    if (acMatch && !acNo) {
      acNo = cleanText(acMatch[1]);
    }
  }

  // Last row of the "Balance" column
  let lastBalanceBDT: number | null = null;
  const headerLine = lines.find(
    (l) => /\bDebit\b/i.test(l.text) && /\bCredit\b/i.test(l.text) && /\bBalance\b/i.test(l.text)
  );
  const totalLineIdx = lines.findIndex((l) => /^Total\s*:/i.test(l.text) || /\bTotal\s*:/i.test(l.text));

  if (headerLine) {
    const balHeaderItem = headerLine.items.find((it) => /\bBalance\b/i.test(it.text));
    const balX = balHeaderItem ? balHeaderItem.x + balHeaderItem.width / 2 : null;
    const headerIdx = lines.indexOf(headerLine);
    const endIdx = totalLineIdx > headerIdx ? totalLineIdx : lines.length;

    const candidateBalances: number[] = [];
    for (let i = headerIdx + 1; i < endIdx; i++) {
      const row = lines[i];
      // Must look like a transaction or Balance B/F row (contains a date or Balance B/F)
      const isTxRow = /^\d{2}\/\d{2}\/\d{4}/.test(row.text) || /Balance\s*B\/F/i.test(row.text);
      if (!isTxRow) continue;

      if (balX !== null) {
        const numItems = row.items.filter((it) => parseIndianNumber(it.text) !== null);
        const balItem = findNearestItemByX(numItems, balX, 75);
        if (balItem) {
          const val = parseIndianNumber(balItem.text);
          if (val !== null) candidateBalances.push(val);
        }
      } else {
        const nums = row.items
          .map((it) => parseIndianNumber(it.text))
          .filter((n): n is number => n !== null);
        if (nums.length > 0) candidateBalances.push(nums[nums.length - 1]);
      }
    }

    if (candidateBalances.length > 0) {
      lastBalanceBDT = candidateBalances[candidateBalances.length - 1];
    }
  }

  return setCachedParse('acctStmt', file, {
    data: {
      accountNumber: acNo,
      acNo,
      branchName,
      customerId,
      lastBalanceBDT,
    },
    rawText,
  });
}

/**
 * 6. CIB Report PDF (#62# CIB date from Date of Inquiry)
 */
export async function parseCibReportPdf(
  file: File,
  onStage?: StageCallback
): Promise<{
  data: CibReportData;
  rawText: string;
}> {
  const cached = getCachedParse<{ data: CibReportData; rawText: string }>('cib', file);
  if (cached) {
    onStage?.('Loaded from cache');
    return cached;
  }

  onStage?.('Reading PDF text...');
  const { lines, rawText } = await extractPdfLines(file);
  onStage?.('Extracting inquiry date...');

  let cibDate = '';

  const inquiryHeaderIdx = lines.findIndex((l) => /Date\s*of\s*Inquiry/i.test(l.text));
  if (inquiryHeaderIdx !== -1) {
    // Check current line or next 2 lines for a date like 17-Nov-2026 or 17/11/2026 or 17.11.2026
    const searchWindow = lines.slice(inquiryHeaderIdx, inquiryHeaderIdx + 3);
    for (const l of searchWindow) {
      const m =
        l.text.match(/(\d{1,2}-[A-Za-z]{3}-\d{4})/) ||
        l.text.match(/(\d{1,2}[./-]\d{1,2}[./-]\d{4})/);
      if (m) {
        cibDate = m[1];
        break;
      }
    }
  }

  if (!cibDate) {
    // Fallback: look for any date near top header
    const firstPage = lines.slice(0, 15).map((l) => l.text).join(' ');
    const m =
      firstPage.match(/(\d{1,2}-[A-Za-z]{3}-\d{4})/) ||
      firstPage.match(/(\d{1,2}\s+[A-Za-z]{3}\s+\d{4})/);
    if (m) cibDate = m[1];
  }

  return setCachedParse('cib', file, {
    data: { cibDate },
    rawText,
  });
}

/**
 * 7. Cost of Fund (.xlsx / .xls via SheetJS, plus .pdf fallback)
 * Extracts "Surplus/Deficit of realization over cost of fund" -> #31#
 */
export async function parseCostOfFundFile(
  file: File,
  onStage?: StageCallback
): Promise<{
  data: CostOfFundData;
  rawText: string;
}> {
  const cached = getCachedParse<{ data: CostOfFundData; rawText: string }>('cof', file);
  if (cached) {
    onStage?.('Loaded from cache');
    return cached;
  }

  const isPdf = file.name.toLowerCase().endsWith('.pdf') || file.type === 'application/pdf';

  if (isPdf) {
    onStage?.('Reading PDF text...');
    const { lines, rawText } = await extractPdfLines(file);
    onStage?.('Extracting cost of fund...');
    let surplusDeficitBDT: number | null = null;
    for (const line of lines) {
      if (/surpluse?\s*\/?\s*deficit\s+of\s+realization\s+over\s+cost\s+of\s+fund/i.test(line.text)) {
        const nums = line.items
          .map((it) => parseIndianNumber(it.text))
          .filter((n): n is number => n !== null);
        if (nums.length > 0) {
          surplusDeficitBDT = nums[nums.length - 1];
          break;
        }
      }
    }
    return setCachedParse('cof', file, { data: { surplusDeficitBDT }, rawText });
  }

  // Excel (.xlsx, .xls, .csv) via SheetJS
  onStage?.('Reading Excel workbook...');
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });
  onStage?.('Scanning sheets...');

  let surplusDeficitBDT: number | null = null;
  const previewRows: string[] = [];

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json<Array<string | number | null>>(sheet, {
      header: 1,
      raw: false,
      defval: '',
    });

    for (const row of rows) {
      const rowText = row.map((c) => String(c ?? '').trim()).filter(Boolean).join(' | ');
      if (rowText) previewRows.push(rowText);

      const hasTargetLabel = row.some((cell) =>
        /surpluse?\s*\/?\s*deficit\s+of\s+realization/i.test(String(cell ?? ''))
      );

      if (hasTargetLabel) {
        // Find the numeric cell in this row (scan from right to left)
        for (let c = row.length - 1; c >= 0; c--) {
          const val = parseIndianNumber(row[c]);
          if (val !== null) {
            surplusDeficitBDT = val;
            break;
          }
        }
        if (surplusDeficitBDT !== null) break;
      }
    }
    if (surplusDeficitBDT !== null) break;
  }

  return setCachedParse('cof', file, {
    data: { surplusDeficitBDT },
    rawText: previewRows.join('\n'),
  });
}

/**
 * Automatically identifies which of the 7 banking documents a file corresponds to
 * for instant batch drag-and-drop / multi-file upload.
 */
export async function detectDocumentKeyFromFile(file: File): Promise<DocumentKey | null> {
  const name = file.name.toLowerCase();

  if (name.endsWith('.xlsx') || name.endsWith('.xls') || name.endsWith('.csv') || /cost.*fund|cof/.test(name)) {
    return 'costOfFund';
  }
  if (/branch.*proposal|proposal/.test(name) || file.type.startsWith('image/')) {
    return 'branchProposal';
  }
  if (/outstanding/.test(name)) {
    return 'outstanding';
  }
  if (/party.*performance|performance/.test(name)) {
    return 'partyPerformance';
  }
  if (/overdue/.test(name)) {
    return 'overdue';
  }
  if (/account.*statement|statement|a\/c/.test(name)) {
    return 'accountStatement';
  }
  if (/cib/.test(name)) {
    return 'cibReport';
  }

  // Fallback: quick text inspection of PDF
  try {
    const { rawText } = await extractPdfLines(file);
    if (/customer\s+outstanding\s+report|princ\s+out/i.test(rawText)) return 'outstanding';
    if (/perform[ea]nce\s+report|recovery\s+amt/i.test(rawText)) return 'partyPerformance';
    if (/statement\s+of\s+overdue/i.test(rawText)) return 'overdue';
    if (/a\/c\s+statement|mudaraba|balance\s+b\/f/i.test(rawText)) return 'accountStatement';
    if (/credit\s+information\s+bureau|cib\s+subject\s+code/i.test(rawText)) return 'cibReport';
    if (/cost\s+of\s+fund|realization\s+over\s+cost/i.test(rawText)) return 'costOfFund';
    // Scanned PDF with negligible extractable text -> likely scanned Branch Proposal
    if (rawText.trim().length < 150) return 'branchProposal';
  } catch {
    // Ignore inspection error
  }

  return null;
}

