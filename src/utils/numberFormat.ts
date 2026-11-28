/**
 * Parses a numeric string tolerating Indian/Bangladeshi comma grouping (e.g., 1,00,000.00),
 * currency prefixes/suffixes (Tk., BDT, /-), whitespace, and parentheses for negatives.
 */
export function parseIndianNumber(raw: string | number | null | undefined): number | null {
  if (raw === null || raw === undefined) return null;
  if (typeof raw === 'number') {
    return Number.isFinite(raw) ? raw : null;
  }

  let s = String(raw).trim();
  if (!s) return null;

  const isNegativeParen = s.startsWith('(') && s.endsWith(')');
  if (isNegativeParen) {
    s = s.slice(1, -1).trim();
  }

  // Strip currency markers, trailing /-, commas, and spaces
  s = s
    .replace(/BDT|Tk\.?|TAKA/gi, '')
    .replace(/\/-/g, '')
    .replace(/,/g, '')
    .replace(/\s+/g, '')
    .trim();

  if (s === '-' || s === '' || /^n\/?a$/i.test(s) || /^nil$/i.test(s)) {
    return null;
  }

  const num = Number(s);
  if (!Number.isFinite(num)) return null;
  return isNegativeParen ? -num : num;
}

/**
 * Converts a BDT amount to Million BDT rounded strictly to 3 decimal places.
 */
export function bdtToMillionStr(bdtAmount: number | null | undefined): string {
  if (bdtAmount === null || bdtAmount === undefined || !Number.isFinite(bdtAmount)) {
    return '';
  }
  const millions = bdtAmount / 1_000_000;
  // Standard half-away-from-zero rounding to 3 decimals
  const rounded = Math.round((millions + Number.EPSILON) * 1000) / 1000;
  const normalized = Object.is(rounded, -0) ? 0 : rounded;
  return normalized.toFixed(3);
}

/**
 * Formats a numeric Million BDT value to 3 decimal places.
 */
export function formatMillionNumber(val: number | null | undefined): string {
  if (val === null || val === undefined || !Number.isFinite(val)) return '';
  const rounded = Math.round((val + Number.EPSILON) * 1000) / 1000;
  const normalized = Object.is(rounded, -0) ? 0 : rounded;
  return normalized.toFixed(3);
}

/**
 * Formats a clean numeric Million BDT string (e.g. "0.50", "0.50 Million") to 3 decimal places,
 * while preserving non-numeric labels (e.g. "N/A", "Nil", "-") untouched.
 */
export function formatMillionString(raw: string | null | undefined): string {
  const s = (raw ?? '').trim();
  if (!s) return '';

  const numOnly = s.match(/^-?\d+(?:,\d+)*(?:\.\d+)?$/);
  if (numOnly) {
    const parsed = parseIndianNumber(s);
    if (parsed !== null) return formatMillionNumber(parsed);
  }

  const numWithSuffix = s.match(/^(-?\d+(?:,\d+)*(?:\.\d+)?)(\s*(?:million|mn|m)\.?)$/i);
  if (numWithSuffix) {
    const parsed = parseIndianNumber(numWithSuffix[1]);
    if (parsed !== null) return `${formatMillionNumber(parsed)}${numWithSuffix[2]}`;
  }

  return s;
}

/**
 * Normalizes a string extracted from a document that may be blank or whitespace.
 */
export function cleanText(val: string | null | undefined): string {
  return (val ?? '').replace(/\s+/g, ' ').trim();
}
