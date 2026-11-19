import JSZip from 'jszip';
import type {
  MemoDocType,
  MemoTemplateRecord,
  ResolvedICMemo,
  TemplatePreviewBlock,
} from '../types/memo';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function directChildren(parent: Element, localName: string): Element[] {
  const out: Element[] = [];
  for (let i = 0; i < parent.childNodes.length; i++) {
    const n = parent.childNodes[i];
    if (n.nodeType === 1 && (n as Element).localName === localName) {
      out.push(n as Element);
    }
  }
  return out;
}

function firstDescendant(parent: Element, localName: string): Element | null {
  for (let i = 0; i < parent.childNodes.length; i++) {
    const n = parent.childNodes[i];
    if (n.nodeType === 1) {
      const el = n as Element;
      if (el.localName === localName) return el;
      const found = firstDescendant(el, localName);
      if (found) return found;
    }
  }
  return null;
}

/** Collect <w:t> nodes belonging directly to this <w:p> (ignoring nested <w:p> inside textboxes) */
function collectDirectTextNodes(pEl: Element): Element[] {
  const nodes: Element[] = [];
  const walk = (el: Element) => {
    for (let i = 0; i < el.childNodes.length; i++) {
      const child = el.childNodes[i];
      if (child.nodeType !== 1) continue;
      const cEl = child as Element;
      if (cEl.localName === 'p') continue; // Do not cross into nested textbox paragraphs
      if (cEl.localName === 't') {
        nodes.push(cEl);
      } else if (cEl.localName !== 'delText' && cEl.localName !== 'instrText') {
        walk(cEl);
      }
    }
  };
  walk(pEl);
  return nodes;
}

function getParagraphPlainText(pEl: Element): string {
  const parts: string[] = [];
  const walk = (el: Element) => {
    for (let i = 0; i < el.childNodes.length; i++) {
      const child = el.childNodes[i];
      if (child.nodeType !== 1) continue;
      const cEl = child as Element;
      if (cEl.localName === 'p') continue;
      if (cEl.localName === 't') {
        parts.push(cEl.textContent ?? '');
      } else if (cEl.localName === 'br' || cEl.localName === 'cr') {
        parts.push('\n');
      } else if (cEl.localName === 'tab') {
        parts.push(' ');
      } else if (cEl.localName !== 'delText' && cEl.localName !== 'instrText') {
        walk(cEl);
      }
    }
  };
  walk(pEl);
  return parts.join('');
}

function getCellPlainText(tcEl: Element): string {
  return directChildren(tcEl, 'p')
    .map((p) => getParagraphPlainText(p).trim())
    .filter(Boolean)
    .join('\n');
}

function setTextNodeValue(tEl: Element, value: string) {
  tEl.textContent = value;
  if (value.startsWith(' ') || value.endsWith(' ') || value.includes('  ')) {
    tEl.setAttribute('xml:space', 'preserve');
  }
}

/**
 * Replaces a substring [start, end) across potentially split <w:t> nodes in a <w:p>,
 * preserving the <w:rPr> of the first overlapping run and handling multi-line values via <w:br/>.
 */
function replaceSpanInParagraph(
  xmlDoc: Document,
  pEl: Element,
  start: number,
  end: number,
  replacement: string
) {
  const tNodes = collectDirectTextNodes(pEl);
  if (tNodes.length === 0) return;

  interface Span {
    node: Element;
    start: number;
    end: number;
    text: string;
  }

  const spans: Span[] = [];
  let cursor = 0;
  for (const node of tNodes) {
    const text = node.textContent ?? '';
    spans.push({ node, start: cursor, end: cursor + text.length, text });
    cursor += text.length;
  }

  const overlapping = spans.filter((s) => s.end > start && s.start < end);
  if (overlapping.length === 0) return;

  const first = overlapping[0];
  const last = overlapping[overlapping.length - 1];
  const prefix = first.text.slice(0, Math.max(0, start - first.start));
  const suffix = last.text.slice(Math.max(0, end - last.start));

  for (let i = 1; i < overlapping.length; i++) {
    setTextNodeValue(overlapping[i].node, '');
  }

  if (!replacement.includes('\n')) {
    setTextNodeValue(first.node, prefix + replacement + suffix);
    return;
  }

  // Multi-line replacement inside a single <w:r>: split by \n and insert <w:br/>
  const parentRun = first.node.parentElement;
  if (!parentRun || parentRun.localName !== 'r') {
    setTextNodeValue(first.node, prefix + replacement.replace(/\n+/g, ' ') + suffix);
    return;
  }

  const lines = replacement.split('\n');
  setTextNodeValue(first.node, prefix + lines[0]);

  let anchor: Node = first.node;
  for (let i = 1; i < lines.length; i++) {
    const br = xmlDoc.createElementNS(W_NS, 'w:br');
    parentRun.insertBefore(br, anchor.nextSibling);
    anchor = br;

    const t = xmlDoc.createElementNS(W_NS, 'w:t');
    const lineVal = i === lines.length - 1 ? lines[i] + suffix : lines[i];
    setTextNodeValue(t, lineVal);
    parentRun.insertBefore(t, anchor.nextSibling);
    anchor = t;
  }
}

/**
 * Replaces the full visible text of a paragraph while keeping <w:pPr> and first run's <w:rPr> intact.
 */
function setParagraphTextPreservingStyle(
  xmlDoc: Document,
  pEl: Element,
  newText: string
) {
  const pPr = directChildren(pEl, 'pPr')[0] ?? null;
  const firstRun = directChildren(pEl, 'r')[0] ?? null;
  const firstRPr = firstRun ? firstDescendant(firstRun, 'rPr') : null;
  const pRPr = pPr ? firstDescendant(pPr, 'rPr') : null;
  const rPrTemplate = firstRPr
    ? (firstRPr.cloneNode(true) as Element)
    : pRPr
      ? (pRPr.cloneNode(true) as Element)
      : null;

  // Keep only pPr, remove existing runs/text content
  const toRemove: ChildNode[] = [];
  for (let i = 0; i < pEl.childNodes.length; i++) {
    const child = pEl.childNodes[i];
    if (child !== pPr) {
      toRemove.push(child);
    }
  }
  for (const child of toRemove) {
    pEl.removeChild(child);
  }

  const rEl = xmlDoc.createElementNS(W_NS, 'w:r');
  if (rPrTemplate) {
    rEl.appendChild(rPrTemplate);
  }

  const lines = newText.split('\n');
  for (let i = 0; i < lines.length; i++) {
    if (i > 0) {
      rEl.appendChild(xmlDoc.createElementNS(W_NS, 'w:br'));
    }
    const tEl = xmlDoc.createElementNS(W_NS, 'w:t');
    setTextNodeValue(tEl, lines[i]);
    rEl.appendChild(tEl);
  }
  pEl.appendChild(rEl);
}

/**
 * Sets text inside a <w:tc> table cell while strictly preserving <w:tcPr>, <w:pPr>, and <w:rPr>.
 */
function setCellTextPreservingFormat(
  xmlDoc: Document,
  tcEl: Element,
  newText: string
) {
  const paras = directChildren(tcEl, 'p');
  if (paras.length === 0) {
    const pEl = xmlDoc.createElementNS(W_NS, 'w:p');
    tcEl.appendChild(pEl);
    setParagraphTextPreservingStyle(xmlDoc, pEl, newText);
    return;
  }

  // Keep first paragraph's exact <w:pPr> and first run's <w:rPr>
  const firstP = paras[0];
  setParagraphTextPreservingStyle(xmlDoc, firstP, newText);

  // Remove any leftover extra paragraphs inside the cell so multi-line replacements don't collide
  for (let i = 1; i < paras.length; i++) {
    tcEl.removeChild(paras[i]);
  }
}

/**
 * Replaces "Label: <old>" inside a paragraph while preserving the bold/underlined label runs
 * and setting the value after ":" using the value run's styling (or non-bold/normal fallback).
 */
function setParagraphValueAfterColon(
  xmlDoc: Document,
  pEl: Element,
  newValue: string
): boolean {
  const full = collectDirectTextNodes(pEl)
    .map((t) => t.textContent ?? '')
    .join('');
  const colonIdx = full.indexOf(':');
  if (colonIdx === -1) return false;

  const afterIdx = colonIdx + 1;
  const tail = full.slice(afterIdx);
  const leadingSpaces = tail.match(/^\s*/)?.[0] ?? ' ';
  const replaceStart = afterIdx + leadingSpaces.length;
  replaceSpanInParagraph(xmlDoc, pEl, replaceStart, full.length, newValue);
  return true;
}

export function buildTemplateTokenMap(memo: ResolvedICMemo): {
  requiredTokens: Record<string, string>;
  optionalTokens: Record<string, string>;
} {
  const ph = memo.placeholders;
  const bp = memo.branchProposal;
  const clientName = bp?.section10Particulars?.nameOfClient?.trim() || ph['#2#'] || '';

  const requiredTokens: Record<string, string> = {};
  for (const [code, val] of Object.entries(ph)) {
    requiredTokens[code] = val && val.trim() !== '' ? val.trim() : code;
  }

  const semanticMap: Record<string, string> = {
    proposalDate: ph['#1#'] || '',
    dated: ph['#1#'] || '',
    memoDate: ph['#1#'] || '',
    accountNo: ph['#2#'] || '',
    accountNumber: ph['#2#'] || '',
    acNo: ph['#2#'] || '',
    clientName,
    nameOfClient: clientName,
    branchName: ph['#3#'] || '',
    branch: ph['#3#'] || '',
    waiverAmount: ph['#4#'] || '',
    waiver: ph['#4#'] || '',
    mode: ph['#5#'] || '',
    modeOfInvestment: ph['#5#'] || '',
    productName: ph['#5#'] || '',
    limitOriginal: ph['#6#'] || '',
    originalLimit: ph['#6#'] || '',
    limit: ph['#6#'] || '',
    principal: ph['#7#'] || '',
    principalOutstanding: ph['#7#'] || '',
    charges: ph['#8#'] || '',
    chargedProfit: ph['#10#'] || '',
    profitOut: ph['#10#'] || '',
    compensation: ph['#11#'] || '',
    unchargedCompensation: ph['#12#'] || '',
    unappliedProfit: ph['#12#'] || '',
    outstanding: ph['#25#'] || '',
    netOutstanding: ph['#25#'] || '',
    totalOutstanding: ph['#25#'] || '',
    settlementAmount: ph['#27#'] || '',
    exitAmount: ph['#27#'] || '',
    settlementPeriod: ph['#28#'] || '',
    settlementDeadline: ph['#28#'] || '',
    costOfFund: ph['#31#'] || '',
    cof: ph['#31#'] || '',
    breakupTotal: ph['#32#'] || '',
    securityMV: ph['#33#'] || '',
    mv: ph['#33#'] || '',
    securityFSV: ph['#34#'] || '',
    fsv: ph['#34#'] || '',
    downPayment: ph['#35#'] || '',
    downPaymentAccount: ph['#36#'] || '',
    totalProfitComp: ph['#42#'] || '',
    maxWaiver: ph['#44#'] || '',
    waiverRecClientBranch: ph['#45#'] || '',
    waiverRecHO: ph['#47#'] || '',
    justification: memo.justificationLines || '',
    natureOfBusiness: ph['#54#'] || '',
    sec2Limit: ph['#56#'] || '',
    validity: ph['#58#'] || '',
    overdue: ph['#59#'] || '',
    recovery: ph['#60#'] || '',
    cibDate: ph['#62#'] || '',
    refNo: ph['#64#'] || '',
    referenceNo: ph['#64#'] || '',
    branchProposalDate: ph['#65#'] || '',
    osDate: ph['#68#'] || '',
    printDate: ph['#68#'] || '',
    cis: ph['#69#'] || '',
    customerId: ph['#69#'] || '',
    sourceOfRepayment: memo.sourceOfRepayment || '',
    branchVisit: memo.branchVisitReport || '',
    branchVisitReport: memo.branchVisitReport || '',
    samdObservation: memo.samdObservation || '',
    samd: memo.samdObservation || '',
  };

  const optionalTokens: Record<string, string> = {};
  for (const [key, val] of Object.entries(semanticMap)) {
    optionalTokens[`#${key}#`] = val;
    optionalTokens[`{{${key}}}`] = val;
    optionalTokens[`[${key}]`] = val;
  }

  return { requiredTokens, optionalTokens };
}

function replaceTokensInParagraph(
  xmlDoc: Document,
  pEl: Element,
  tokenMap: Record<string, string>
): boolean {
  let replacedAny = false;
  const keys = Object.keys(tokenMap).sort((a, b) => b.length - a.length);
  if (keys.length === 0) return false;

  // Iterative scan so multiple tokens in one paragraph are all replaced cleanly
  let safety = 0;
  while (safety < 40) {
    safety++;
    const fullText = collectDirectTextNodes(pEl)
      .map((t) => t.textContent ?? '')
      .join('');
    if (!fullText) break;

    let bestIdx = -1;
    let bestKey = '';

    for (const key of keys) {
      const idx = fullText.indexOf(key);
      if (idx !== -1 && (bestIdx === -1 || idx < bestIdx)) {
        // Avoid infinite loop if replacement equals key (e.g., unfilled #code#)
        if (tokenMap[key] !== key) {
          bestIdx = idx;
          bestKey = key;
        }
      }
    }

    if (bestIdx === -1) break;
    replaceSpanInParagraph(xmlDoc, pEl, bestIdx, bestIdx + bestKey.length, tokenMap[bestKey]);
    replacedAny = true;
  }

  return replacedAny;
}

function hasAnyKnownToken(text: string, tokenKeys: string[]): boolean {
  return tokenKeys.some((k) => text.includes(k));
}

function nonEmpty(v: string | undefined): string | null {
  if (!v) return null;
  const trimmed = v.trim();
  return trimmed ? trimmed : null;
}

/**
 * Structural & label-based population for uploaded .docx templates that do not use explicit #code# tokens
 * (or for untokenized sections/tables in a standard Bank Memo template).
 * Strictly preserves all tables, borders, column widths, fonts, sizes, margins, orientation, headers, and footers.
 */
function applyStructuralPopulations(
  xmlDoc: Document,
  memo: ResolvedICMemo,
  paragraphsWithTokens: Set<Element>
) {
  const ph = memo.placeholders;
  const bp = memo.branchProposal;
  const clientName = nonEmpty(bp?.section10Particulars?.nameOfClient) || nonEmpty(ph['#2#']);

  const bodyEl = firstDescendant(xmlDoc.documentElement, 'body');
  if (!bodyEl) return;

  // 1. Process direct body paragraphs
  const bodyParas = directChildren(bodyEl, 'p');
  for (let i = 0; i < bodyParas.length; i++) {
    const pEl = bodyParas[i];
    if (paragraphsWithTokens.has(pEl)) continue;

    const text = getParagraphPlainText(pEl).trim();
    if (!text) continue;

    // 1. Nature of Business:
    if (/^1\.\s*Nature of Business\s*:/i.test(text) && nonEmpty(ph['#54#'])) {
      setParagraphValueAfterColon(xmlDoc, pEl, ph['#54#']);
      continue;
    }

    // 2. Liability Position ... as on ...:
    if (/^2\.\s*Liability Position.*as on\b/i.test(text) && nonEmpty(ph['#68#'])) {
      const m = text.match(/as on\s+([^:]+):/i);
      if (m) {
        const start = text.indexOf(m[1], text.toLowerCase().indexOf('as on'));
        if (start !== -1) {
          replaceSpanInParagraph(xmlDoc, pEl, start, start + m[1].length, ph['#68#']);
        }
      }
      continue;
    }

    // 4. CIB Status: Report dated ... shows:
    if (/^4\.\s*CIB Status\s*:/i.test(text) && nonEmpty(ph['#62#'])) {
      const m = text.match(/Report dated\s+(.+?)\s+shows/i);
      if (m) {
        const start = text.indexOf(m[1]);
        if (start !== -1) {
          replaceSpanInParagraph(xmlDoc, pEl, start, start + m[1].length, ph['#62#']);
        }
      }
      continue;
    }

    // 8. Source of Repayment:
    if (/^8\.\s*Source of Repayment\s*:/i.test(text) && nonEmpty(memo.sourceOfRepayment)) {
      setParagraphValueAfterColon(xmlDoc, pEl, memo.sourceOfRepayment);
      continue;
    }

    // 12. Branch Visit Report: (ONLY populate if generated/edited)
    if (/^12\.\s*Branch Visit Report\s*:/i.test(text) && nonEmpty(memo.branchVisitReport)) {
      setParagraphValueAfterColon(xmlDoc, pEl, memo.branchVisitReport);
      continue;
    }

    // 13. Recommendation of The Branch:
    if (/^13\.\s*Recommendation of The Branch\s*:/i.test(text)) {
      let updated = text;
      if (nonEmpty(ph['#64#'])) {
        updated = updated.replace(/(Proposal\s*Ref\s*:\s*)(.+?)(\s+dated\b)/i, `$1${ph['#64#']}$3`);
      }
      if (nonEmpty(ph['#65#'])) {
        updated = updated.replace(/(\bdated\s+)(.+)$/i, `$1${ph['#65#']}`);
      }
      if (updated !== text) {
        const colonIdx = text.indexOf(':');
        if (colonIdx !== -1) {
          setParagraphValueAfterColon(xmlDoc, pEl, updated.slice(colonIdx + 1).trim());
        }
      }
      continue;
    }

    // Subject: ... waiver of Tk. ...
    if (/^Subject\s*:/i.test(text) && nonEmpty(ph['#4#'])) {
      const m = text.match(/waiver of Tk\.?\s*([^\s,]+(?:\s*million)?)/i);
      if (m) {
        const start = text.indexOf(m[1], text.toLowerCase().indexOf('waiver of tk'));
        if (start !== -1) {
          replaceSpanInParagraph(xmlDoc, pEl, start, start + m[1].length, ph['#4#']);
        }
      }
      continue;
    }

    // Justification of The Proposed Facility: -> update following "a." line if justificationLines exists
    if (/^Justification of The Proposed Facility\s*:?/i.test(text) && nonEmpty(memo.justificationLines)) {
      const nextP = bodyParas[i + 1];
      if (nextP && !paragraphsWithTokens.has(nextP)) {
        const nextTxt = getParagraphPlainText(nextP).trim();
        if (/^a\./i.test(nextTxt)) {
          const cleanLines = memo.justificationLines
            .split('\n')
            .map((l) => l.trim())
            .filter(Boolean)
            .join('\n');
          replaceSpanInParagraph(xmlDoc, nextP, 2, nextTxt.length, ` ${cleanLines}`);
        }
      }
      continue;
    }

    // 15. Observation and Recommendation of SAMD: (ONLY populate if generated/edited)
    if (/^15\.\s*Observation and Recommendation of SAMD\s*:?/i.test(text) && nonEmpty(memo.samdObservation)) {
      const obsLines = memo.samdObservation
        .split('\n')
        .map((l) => l.trim())
        .filter(Boolean);
      if (obsLines.length > 0) {
        // Check existing sub-paragraphs between section 15 and section 16
        const existingSub: Element[] = [];
        for (let j = i + 1; j < bodyParas.length; j++) {
          const subText = getParagraphPlainText(bodyParas[j]).trim();
          if (/^16\./.test(subText) || /^17\./.test(subText)) break;
          // Stop if a table appears immediately after in document order
          let node: ChildNode | null = bodyParas[j].previousSibling;
          let hitTable = false;
          while (node && node !== pEl) {
            if (node.nodeType === 1 && (node as Element).localName === 'tbl') {
              hitTable = true;
              break;
            }
            node = node.previousSibling;
          }
          if (hitTable) break;
          existingSub.push(bodyParas[j]);
        }

        const templateSub = existingSub[0] ?? null;
        if (existingSub.length > 0) {
          // Put all lines into first sub-paragraph (or clone per line if multiple)
          const firstSub = existingSub[0];
          setParagraphTextPreservingStyle(xmlDoc, firstSub, obsLines[0]);
          let insertAfter: Element = firstSub;
          for (let k = 1; k < obsLines.length; k++) {
            if (k < existingSub.length) {
              setParagraphTextPreservingStyle(xmlDoc, existingSub[k], obsLines[k]);
              insertAfter = existingSub[k];
            } else {
              const cloned = templateSub!.cloneNode(true) as Element;
              setParagraphTextPreservingStyle(xmlDoc, cloned, obsLines[k]);
              insertAfter.parentNode?.insertBefore(cloned, insertAfter.nextSibling);
              insertAfter = cloned;
            }
          }
          // Clear any extra old lines if template had more than obsLines.length
          for (let k = obsLines.length; k < existingSub.length; k++) {
            existingSub[k].parentNode?.removeChild(existingSub[k]);
          }
        } else {
          // No existing sub-paragraph after heading: clone heading <w:p> (stripping bold/underline)
          let insertAfter: Element = pEl;
          for (const line of obsLines) {
            const cloned = xmlDoc.createElementNS(W_NS, 'w:p');
            const pPr = xmlDoc.createElementNS(W_NS, 'w:pPr');
            const ind = xmlDoc.createElementNS(W_NS, 'w:ind');
            ind.setAttribute('w:left', '180');
            pPr.appendChild(ind);
            const sp = xmlDoc.createElementNS(W_NS, 'w:spacing');
            sp.setAttribute('w:before', '15');
            sp.setAttribute('w:after', '15');
            pPr.appendChild(sp);
            cloned.appendChild(pPr);

            // Inherit font & size from heading run if available
            const headingRun = directChildren(pEl, 'r')[0];
            const headingRPr = headingRun ? firstDescendant(headingRun, 'rPr') : null;
            const rEl = xmlDoc.createElementNS(W_NS, 'w:r');
            if (headingRPr) {
              const cleanRPr = headingRPr.cloneNode(true) as Element;
              for (const tag of ['b', 'bCs', 'u']) {
                const el = firstDescendant(cleanRPr, tag);
                if (el) el.parentNode?.removeChild(el);
              }
              rEl.appendChild(cleanRPr);
            }
            const tEl = xmlDoc.createElementNS(W_NS, 'w:t');
            setTextNodeValue(tEl, line);
            rEl.appendChild(tEl);
            cloned.appendChild(rEl);

            insertAfter.parentNode?.insertBefore(cloned, insertAfter.nextSibling);
            insertAfter = cloned;
          }
        }
      }
    }
  }

  // 2. Process all tables in document
  const tables = Array.from(xmlDoc.getElementsByTagNameNS(W_NS, 'tbl'));
  for (const tbl of tables) {
    const trs = directChildren(tbl, 'tr');
    if (trs.length === 0) continue;

    const matrix = trs.map((tr) =>
      directChildren(tr, 'tc').map((tc) => ({
        tc,
        text: getCellPlainText(tc),
        hasTokens: directChildren(tc, 'p').some((p) => paragraphsWithTokens.has(p)),
      }))
    );

    const tableText = matrix.flat().map((c) => c.text).join(' | ');

    // A. Metadata Table (Dated / Memo no / To / From)
    if (/\bDated\b/i.test(tableText) && /\bMemo\s*no\b/i.test(tableText)) {
      for (const row of matrix) {
        const label = row[0]?.text.trim() ?? '';
        const valCell = row[row.length - 1];
        if (!valCell || valCell.hasTokens) continue;
        if (/^Dated$/i.test(label) && nonEmpty(ph['#1#'])) {
          setCellTextPreservingFormat(xmlDoc, valCell.tc, ph['#1#']);
        }
      }
      continue;
    }

    // B. A/C & Branch + Subject Bar
    if (/\bA\/C\s*:/i.test(tableText) && /\bBranch\s*:/i.test(tableText)) {
      for (const row of matrix) {
        for (const c of row) {
          if (c.hasTokens) continue;
          if (/\bA\/C\s*:/i.test(c.text) && nonEmpty(ph['#2#'])) {
            const cis = nonEmpty(ph['#69#']) || '#69#';
            setCellTextPreservingFormat(xmlDoc, c.tc, `A/C: ${ph['#2#']} (CIS: ${cis})`);
          } else if (/\bBranch\s*:/i.test(c.text) && nonEmpty(ph['#3#'])) {
            setCellTextPreservingFormat(xmlDoc, c.tc, `Branch: ${ph['#3#']}`);
          } else if (/^Subject\s*:/i.test(c.text) && nonEmpty(ph['#4#'])) {
            const updated = c.text.replace(
              /(waiver of Tk\.?\s*)([^\s,]+)/i,
              `$1${ph['#4#']}`
            );
            if (updated !== c.text) {
              setCellTextPreservingFormat(xmlDoc, c.tc, updated);
            }
          }
        }
      }
      continue;
    }

    // C. Main Settlement Table (Particulars + Grand-Total + Principal)
    if (/\bPrincipal\b/i.test(tableText) && /\bGrand[-\s]*Total\b/i.test(tableText)) {
      for (let r = 0; r < matrix.length; r++) {
        const row = matrix[r];
        const joined = row.map((c) => c.text).join(' | ');

        // Header row "O/S as on ..."
        for (const c of row) {
          if (!c.hasTokens && /O\/S\s+as\s+on\b/i.test(c.text) && nonEmpty(ph['#68#'])) {
            const updated = c.text.replace(/(O\/S\s+as\s+on\s+).*/i, `$1${ph['#68#']}`);
            setCellTextPreservingFormat(xmlDoc, c.tc, updated);
          }
        }

        // First data row containing Principal (also contains Mode, Limit Original, RSD, Remarks via rowSpan)
        const princCellIdx = row.findIndex((c) => /^Principal$/i.test(c.text.trim()));
        if (princCellIdx !== -1) {
          if (princCellIdx >= 3) {
            if (!row[0].hasTokens && nonEmpty(ph['#5#'])) {
              setCellTextPreservingFormat(xmlDoc, row[0].tc, ph['#5#']);
            }
            if (!row[1].hasTokens && nonEmpty(ph['#6#'])) {
              setCellTextPreservingFormat(xmlDoc, row[1].tc, ph['#6#']);
            }
            if (!row[2].hasTokens && nonEmpty(bp?.limitRSD)) {
              setCellTextPreservingFormat(xmlDoc, row[2].tc, bp!.limitRSD);
            }
          }
          const osCell = row[princCellIdx + 1];
          if (osCell && !osCell.hasTokens && nonEmpty(ph['#7#'])) {
            setCellTextPreservingFormat(xmlDoc, osCell.tc, ph['#7#']);
          }
          const remarksCell = row[row.length - 1];
          if (
            remarksCell &&
            !remarksCell.hasTokens &&
            /Settlement\s*Deadline|Break\s*up|Security\s*Value/i.test(remarksCell.text)
          ) {
            const v = (code: string) => nonEmpty(ph[code]) || code;
            const remarks = `1)Settlement Deadline:\n${v('#28#')}\n\n2) Break up\nPrincipal: ${v('#7#')}\nCharges: ${v('#8#')}\nCOF: ${v('#31#')}\nTotal: ${v('#32#')}\n\n3) Security Value:\nMV: ${v('#33#')}\nFSV: ${v('#34#')}`;
            setCellTextPreservingFormat(xmlDoc, remarksCell.tc, remarks);
          }
          continue;
        }

        if (/^Charges\b/i.test(row[0]?.text.trim() ?? '') && !row[1]?.hasTokens && nonEmpty(ph['#8#'])) {
          setCellTextPreservingFormat(xmlDoc, row[1].tc, ph['#8#']);
          continue;
        }
        if (/^Charged Profit$/i.test(row[0]?.text.trim() ?? '') && !row[1]?.hasTokens && nonEmpty(ph['#10#'])) {
          setCellTextPreservingFormat(xmlDoc, row[1].tc, ph['#10#']);
          continue;
        }
        if (/^Compensation$/i.test(row[0]?.text.trim() ?? '') && !row[1]?.hasTokens && nonEmpty(ph['#11#'])) {
          setCellTextPreservingFormat(xmlDoc, row[1].tc, ph['#11#']);
          continue;
        }
        if (/^Uncharged Compensation$/i.test(row[0]?.text.trim() ?? '') && !row[1]?.hasTokens && nonEmpty(ph['#12#'])) {
          setCellTextPreservingFormat(xmlDoc, row[1].tc, ph['#12#']);
          continue;
        }
        if (/Grand[-\s]*Total/i.test(joined)) {
          if (row[1] && !row[1].hasTokens && nonEmpty(ph['#25#'])) {
            setCellTextPreservingFormat(xmlDoc, row[1].tc, ph['#25#']);
          }
          if (row[2] && !row[2].hasTokens && nonEmpty(ph['#4#'])) {
            setCellTextPreservingFormat(xmlDoc, row[2].tc, ph['#4#']);
          }
          if (row[3] && !row[3].hasTokens && nonEmpty(ph['#27#'])) {
            setCellTextPreservingFormat(xmlDoc, row[3].tc, ph['#27#']);
          }
        }
      }
      continue;
    }

    // D. Down Payment Note Bar
    if (/Note:\s*Information Regarding Down payment/i.test(tableText) && matrix[0]?.[0]) {
      const c = matrix[0][0];
      if (!c.hasTokens && (nonEmpty(ph['#35#']) || nonEmpty(ph['#36#']) || nonEmpty(ph['#68#']))) {
        const dp = nonEmpty(ph['#35#']) || '#35#';
        const ac = nonEmpty(ph['#36#']) || '#36#';
        const dt = nonEmpty(ph['#68#']) || '#68#';
        setCellTextPreservingFormat(
          xmlDoc,
          c.tc,
          `Note: Information Regarding Down payment (deposited Tk.${dp} Million in A/C: ${ac} on ${dt}.)`
        );
      }
      continue;
    }

    // E. Key Points Table
    if (/Cost of Fund/i.test(tableText) && /Maximum Waiver/i.test(tableText)) {
      for (const row of matrix) {
        for (let ci = 0; ci < row.length - 1; ci++) {
          const label = row[ci].text;
          const next = row[ci + 1];
          if (!next || next.hasTokens) continue;
          if (/Total\s*\[Profit/i.test(label) && nonEmpty(ph['#42#'])) {
            setCellTextPreservingFormat(xmlDoc, next.tc, `: ${ph['#42#']}`);
          } else if (/^Cost of Fund$/i.test(label) && nonEmpty(ph['#31#'])) {
            setCellTextPreservingFormat(xmlDoc, next.tc, `: ${ph['#31#']}`);
          } else if (/Maximum Waiver/i.test(label) && nonEmpty(ph['#44#'])) {
            setCellTextPreservingFormat(xmlDoc, next.tc, `: ${ph['#44#']}`);
          } else if (/^(Client|Branch)$/i.test(label) && nonEmpty(ph['#45#'])) {
            setCellTextPreservingFormat(xmlDoc, next.tc, `: ${ph['#45#']}`);
          } else if (/^(SAMD,\s*HO|HOIC|MEC|IC\s*&\s*CD)$/i.test(label) && nonEmpty(ph['#47#'])) {
            setCellTextPreservingFormat(xmlDoc, next.tc, `: ${ph['#47#']}`);
          }
        }
      }
      continue;
    }

    // F. Section 1 Nature of Business single-cell box
    if (/^1\.\s*Nature of Business\s*:/i.test(tableText) && matrix[0]?.[0] && !matrix[0][0].hasTokens) {
      if (nonEmpty(ph['#54#'])) {
        setCellTextPreservingFormat(xmlDoc, matrix[0][0].tc, `1. Nature of Business: ${ph['#54#']}`);
      }
      continue;
    }

    // G. Section 2 Liability Position Table
    if (/Name of the Concern/i.test(tableText) && /Mode of Investment/i.test(tableText)) {
      for (const row of matrix) {
        if (row.length >= 8 && !row[0].hasTokens && /Funded/i.test(row[1]?.text ?? '')) {
          if (nonEmpty(ph['#2#'])) setCellTextPreservingFormat(xmlDoc, row[0].tc, ph['#2#']);
          if (nonEmpty(ph['#55#'])) setCellTextPreservingFormat(xmlDoc, row[1].tc, `${ph['#55#']}\nFunded`);
          if (nonEmpty(ph['#56#'])) setCellTextPreservingFormat(xmlDoc, row[2].tc, ph['#56#']);
          if (nonEmpty(ph['#25#'])) setCellTextPreservingFormat(xmlDoc, row[3].tc, ph['#25#']);
          if (nonEmpty(ph['#58#'])) setCellTextPreservingFormat(xmlDoc, row[4].tc, ph['#58#']);
          if (nonEmpty(ph['#59#'])) setCellTextPreservingFormat(xmlDoc, row[5].tc, ph['#59#']);
          if (nonEmpty(ph['#60#'])) setCellTextPreservingFormat(xmlDoc, row[6].tc, ph['#60#']);
          if (nonEmpty(bp?.sec2Remarks)) setCellTextPreservingFormat(xmlDoc, row[7].tc, bp!.sec2Remarks);
        } else if (/^Total$/i.test(row[0]?.text ?? '') && row.length >= 7) {
          if (!row[1].hasTokens && nonEmpty(ph['#56#'])) setCellTextPreservingFormat(xmlDoc, row[1].tc, ph['#56#']);
          if (!row[2].hasTokens && nonEmpty(ph['#25#'])) setCellTextPreservingFormat(xmlDoc, row[2].tc, ph['#25#']);
          if (!row[3].hasTokens && nonEmpty(ph['#58#'])) setCellTextPreservingFormat(xmlDoc, row[3].tc, ph['#58#']);
          if (!row[4].hasTokens && nonEmpty(ph['#59#'])) setCellTextPreservingFormat(xmlDoc, row[4].tc, ph['#59#']);
          if (!row[5].hasTokens && nonEmpty(ph['#60#'])) setCellTextPreservingFormat(xmlDoc, row[5].tc, ph['#60#']);
        }
      }
      continue;
    }

    // H. Section 16 Particulars of the facility Table
    if (/Settlement\/Exit amount/i.test(tableText) && /Down Payment/i.test(tableText)) {
      for (const row of matrix) {
        if (row.length < 2) continue;
        const label = row[row.length - 2].text;
        const val = row[row.length - 1];
        if (val.hasTokens) continue;
        const dt = nonEmpty(ph['#68#']) || '#68#';
        if (/^Outstanding$/i.test(label) && nonEmpty(ph['#25#'])) {
          setCellTextPreservingFormat(xmlDoc, val.tc, `Tk. ${ph['#25#']} million as on ${dt}.`);
        } else if (/^Down Payment$/i.test(label) && nonEmpty(ph['#35#'])) {
          const ac = nonEmpty(ph['#36#']) || '#36#';
          setCellTextPreservingFormat(xmlDoc, val.tc, `Down-payment of Tk. ${ph['#35#']} Million is kept in ${ac} A/C ${ac}`);
        } else if (/^Waiver$/i.test(label) && nonEmpty(ph['#4#'])) {
          setCellTextPreservingFormat(xmlDoc, val.tc, `Tk. ${ph['#4#']} million as on ${dt}`);
        } else if (/Settlement\/Exit amount/i.test(label) && nonEmpty(ph['#27#'])) {
          setCellTextPreservingFormat(xmlDoc, val.tc, `Tk. ${ph['#27#']} Million as on ${dt}`);
        } else if (/^Settlement period$/i.test(label) && nonEmpty(ph['#28#'])) {
          setCellTextPreservingFormat(xmlDoc, val.tc, ph['#28#']);
        }
      }
      continue;
    }

    // I. Generic Key-Value rows (2-column [Label, Value] or 3/4-column [SL, Label, (:), Value])
    for (const row of matrix) {
      if (row.length < 2 || row.length > 4) continue;
      const valCell = row[row.length - 1];
      if (!valCell || valCell.hasTokens) continue;

      // Determine label cell (skip ":" separator cell)
      let labelCell = row[row.length - 2];
      if (labelCell.text.trim() === ':' && row.length >= 3) {
        labelCell = row[row.length - 3];
      }
      const label = labelCell.text.trim();
      if (!label) continue;

      const sp = bp?.section10Particulars;
      const trySet = (val: string | null) => {
        if (val !== null) setCellTextPreservingFormat(xmlDoc, valCell.tc, val);
      };

      if (/^Name of the Client$/i.test(label)) trySet(nonEmpty(sp?.nameOfClient) || clientName);
      else if (/^Unique Customer ID$/i.test(label)) trySet(nonEmpty(sp?.uniqueCustomerId) || nonEmpty(ph['#69#']));
      else if (/^AWCD\/MSD Account/i.test(label)) trySet(nonEmpty(sp?.awcdMsdAccount));
      else if (/^Business Type\/Nature$/i.test(label)) trySet(nonEmpty(sp?.businessTypeNature) || nonEmpty(ph['#54#']));
      else if (/^Business Management Type$/i.test(label)) trySet(nonEmpty(sp?.businessManagementType));
      else if (/^Client Type$/i.test(label)) trySet(nonEmpty(sp?.clientType));
      else if (/^Key Person/i.test(label)) trySet(nonEmpty(sp?.keyPersonPhone));
      else if (/^Address$/i.test(label)) trySet(nonEmpty(sp?.address));
      else if (/^Office\/Business Address$/i.test(label)) trySet(nonEmpty(sp?.officeBusinessAddress));
      else if (/^Factory Address$/i.test(label)) trySet(nonEmpty(sp?.factoryAddress));
      else if (/^Purpose of Investment$/i.test(label)) trySet(nonEmpty(sp?.purposeOfInvestment));
      else if (/^Limit of Investment$/i.test(label)) trySet(nonEmpty(sp?.limitOfInvestment) || nonEmpty(ph['#6#']));
      else if (/^Mode of Investment$/i.test(label)) trySet(nonEmpty(sp?.modeOfInvestment) || nonEmpty(ph['#5#']));
      else if (/^Last Visit Date$/i.test(label)) trySet(nonEmpty(sp?.lastVisitDate));
      else if (/^Stock Position$/i.test(label)) trySet(nonEmpty(sp?.stockPosition));
      else if (/^Fixed Asset$/i.test(label)) trySet(nonEmpty(sp?.fixedAsset));
      else if (/^Business Net Worth$/i.test(label)) trySet(nonEmpty(sp?.businessNetWorth));
      else if (/^Insurance Coverage$/i.test(label)) trySet(nonEmpty(sp?.insuranceCoverage));
      else if (/^Trade License/i.test(label)) trySet(nonEmpty(sp?.tradeLicenseValidity));
      else if (/^Other Regulatory Documents$/i.test(label)) trySet(nonEmpty(sp?.otherRegulatoryDocuments));
      else if (/^Date of Original Investment$/i.test(label)) trySet(nonEmpty(sp?.dateOfOriginalInvestment));
      else if (/Disbursement Manager$/i.test(label)) trySet(nonEmpty(sp?.firstDisbursementManager));
      else if (/First Date of Classification/i.test(label)) trySet(nonEmpty(sp?.firstDateOfClassification));
      else if (/^(Account\s*No\.?|A\/C\s*No\.?)$/i.test(label)) trySet(nonEmpty(ph['#2#']));
      else if (/^(Branch|Branch\s*Name)$/i.test(label)) trySet(nonEmpty(ph['#3#']));
      else if (/^(Proposal\s*Date|Sanction\s*Date|Memo\s*Date)$/i.test(label)) trySet(nonEmpty(ph['#1#']));
      else if (/^(Proposed\s*Waiver|Waiver\s*Amount)$/i.test(label)) trySet(nonEmpty(ph['#4#']));
      else if (/^(Settlement\s*Amount|Exit\s*Amount)$/i.test(label)) trySet(nonEmpty(ph['#27#']));
      else if (/^(Settlement\s*Period|Settlement\s*Deadline)$/i.test(label)) trySet(nonEmpty(ph['#28#']));
      else if (/^Source of Repayment$/i.test(label)) trySet(nonEmpty(memo.sourceOfRepayment));
      else if (/^Branch Visit Report$/i.test(label)) trySet(nonEmpty(memo.branchVisitReport));
      else if (/^Observation and Recommendation of SAMD$/i.test(label)) trySet(nonEmpty(memo.samdObservation));
      else if (/^Justification of The Proposed Facility$/i.test(label)) trySet(nonEmpty(memo.justificationLines));
    }
  }
}

export async function inspectDocxTemplate(
  file: File,
  docType: MemoDocType
): Promise<MemoTemplateRecord> {
  const lower = file.name.toLowerCase();
  if (!lower.endsWith('.docx')) {
    throw new Error('Please upload a Microsoft Word (.docx) file for exact template layout preservation.');
  }

  const buffer = await file.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  const zip = await JSZip.loadAsync(bytes);

  const docFile = zip.file('word/document.xml');
  if (!docFile) {
    throw new Error('Invalid .docx file: missing word/document.xml.');
  }

  const parser = new DOMParser();
  const tokenRegex = /(#\d+#|#[a-zA-Z][a-zA-Z0-9_]*#|\{\{[a-zA-Z0-9_.]+\}\})/g;
  const detectedSet = new Set<string>();
  const previewSections: TemplatePreviewBlock[] = [];

  const xmlPaths = Object.keys(zip.files).filter((p) =>
    /^word\/(document|header\d*|footer\d*)\.xml$/i.test(p)
  );

  for (const xp of xmlPaths) {
    const xmlStr = await zip.files[xp].async('string');
    const xmlDoc = parser.parseFromString(xmlStr, 'application/xml');
    const pEls = Array.from(xmlDoc.getElementsByTagNameNS(W_NS, 'p'));
    for (const pEl of pEls) {
      const text = getParagraphPlainText(pEl);
      let m: RegExpExecArray | null;
      while ((m = tokenRegex.exec(text)) !== null) {
        detectedSet.add(m[1]);
      }
      tokenRegex.lastIndex = 0;
    }

    if (xp === 'word/document.xml') {
      const bodyEl = firstDescendant(xmlDoc.documentElement, 'body') ?? xmlDoc.documentElement;
      for (let i = 0; i < bodyEl.childNodes.length; i++) {
        const node = bodyEl.childNodes[i];
        if (node.nodeType !== 1) continue;
        const el = node as Element;
        if (el.localName === 'p') {
          const t = getParagraphPlainText(el).trim();
          if (t) previewSections.push({ type: 'paragraph', text: t });
        } else if (el.localName === 'tbl') {
          const rows = directChildren(el, 'tr').map((tr) =>
            directChildren(tr, 'tc').map((tc) => getCellPlainText(tc))
          );
          if (rows.length > 0) {
            previewSections.push({ type: 'table', rows });
          }
        }
      }
    }
  }

  return {
    docType,
    fileName: file.name,
    fileSize: file.size,
    uploadedAt: new Date().toISOString(),
    bytes,
    detectedTokens: Array.from(detectedSet),
    previewSections,
  };
}

/**
 * Populates an uploaded .docx template strictly preserving 100% of its XML structure,
 * margins, orientation, fonts, sizes, spacing, borders, column widths, headers, and footers.
 */
export async function populateDocxTemplate(
  templateBytes: Uint8Array,
  memo: ResolvedICMemo,
  _docType: MemoDocType
): Promise<Blob> {
  const zip = await JSZip.loadAsync(templateBytes);
  const { requiredTokens, optionalTokens } = buildTemplateTokenMap(memo);
  const allTokenMap: Record<string, string> = {
    ...optionalTokens,
    ...requiredTokens,
  };
  const allTokenKeys = Object.keys(allTokenMap);

  const parser = new DOMParser();
  const serializer = new XMLSerializer();

  const xmlPaths = Object.keys(zip.files).filter((p) =>
    /^word\/(document|header\d*|footer\d*)\.xml$/i.test(p)
  );

  for (const xp of xmlPaths) {
    const origXmlStr = await zip.files[xp].async('string');
    const xmlDeclMatch = origXmlStr.match(/^<\?xml[^?]*\?>\s*/i);
    const xmlDoc = parser.parseFromString(origXmlStr, 'application/xml');

    const paragraphsWithTokens = new Set<Element>();
    const allParagraphs = Array.from(xmlDoc.getElementsByTagNameNS(W_NS, 'p'));

    for (const pEl of allParagraphs) {
      const pText = getParagraphPlainText(pEl);
      if (hasAnyKnownToken(pText, allTokenKeys)) {
        const replaced = replaceTokensInParagraph(xmlDoc, pEl, allTokenMap);
        if (replaced) {
          paragraphsWithTokens.add(pEl);
        }
      }
    }

    if (xp === 'word/document.xml') {
      applyStructuralPopulations(xmlDoc, memo, paragraphsWithTokens);
    }

    let serialized = serializer.serializeToString(xmlDoc);
    if (xmlDeclMatch && !serialized.startsWith('<?xml')) {
      serialized = xmlDeclMatch[0] + serialized;
    }
    zip.file(xp, serialized);
  }

  return await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

function applyPreviewReplacements(text: string, tokenMap: Record<string, string>): string {
  let out = text;
  const keys = Object.keys(tokenMap).sort((a, b) => b.length - a.length);
  for (const key of keys) {
    if (out.includes(key) && tokenMap[key] !== key) {
      out = out.split(key).join(tokenMap[key]);
    }
  }
  return out;
}

/**
 * Generates live preview blocks for an uploaded .docx template using the current synchronized memo state.
 */
export function renderTemplatePreviewBlocks(
  template: MemoTemplateRecord,
  memo: ResolvedICMemo
): TemplatePreviewBlock[] {
  const { requiredTokens, optionalTokens } = buildTemplateTokenMap(memo);
  const allTokenMap = { ...optionalTokens, ...requiredTokens };
  const ph = memo.placeholders;

  return template.previewSections.map((block) => {
    if (block.type === 'paragraph') {
      let text = applyPreviewReplacements(block.text || '', allTokenMap);
      if (/^1\.\s*Nature of Business\s*:/i.test(text) && nonEmpty(ph['#54#']) && !block.text?.includes('#54#')) {
        text = `1. Nature of Business: ${ph['#54#']}`;
      } else if (/^8\.\s*Source of Repayment\s*:/i.test(text) && nonEmpty(memo.sourceOfRepayment)) {
        text = `8. Source of Repayment: ${memo.sourceOfRepayment}`;
      } else if (/^12\.\s*Branch Visit Report\s*:/i.test(text) && nonEmpty(memo.branchVisitReport)) {
        text = `12. Branch Visit Report: ${memo.branchVisitReport}`;
      } else if (/^15\.\s*Observation and Recommendation of SAMD\s*:?/i.test(text) && nonEmpty(memo.samdObservation)) {
        text = `15. Observation and Recommendation of SAMD:\n${memo.samdObservation}`;
      }
      return { type: 'paragraph', text };
    }

    const rows = (block.rows || []).map((r) =>
      r.map((cellText) => applyPreviewReplacements(cellText, allTokenMap))
    );
    return { type: 'table', rows };
  });
}
