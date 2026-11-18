import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  ImageRun,
  type ITableCellBorders,
  Packer,
  PageNumber,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  type TableVerticalAlign,
  TextRun,
  UnderlineType,
  VerticalAlign,
  WidthType,
} from 'docx';
import type { ResolvedICMemo } from '../types/memo';

const FONT = 'Times New Roman';
const SZ_BODY = 20; // 10pt
const SZ_SMALL = 18; // 9pt
const SZ_TITLE = 24; // 12pt

const THIN_BORDER = {
  style: BorderStyle.SINGLE,
  size: 4,
  color: '000000',
};

const DOTTED_BORDER = {
  style: BorderStyle.DOTTED,
  size: 4,
  color: '444444',
};

const NO_BORDER = {
  style: BorderStyle.NONE,
  size: 0,
  color: 'FFFFFF',
};

const ALL_THIN_BORDERS: ITableCellBorders = {
  top: THIN_BORDER,
  bottom: THIN_BORDER,
  left: THIN_BORDER,
  right: THIN_BORDER,
};

const ALL_DOTTED_BORDERS: ITableCellBorders = {
  top: DOTTED_BORDER,
  bottom: DOTTED_BORDER,
  left: DOTTED_BORDER,
  right: DOTTED_BORDER,
};

const ALL_NO_BORDERS: ITableCellBorders = {
  top: NO_BORDER,
  bottom: NO_BORDER,
  left: NO_BORDER,
  right: NO_BORDER,
};

function txt(
  text: string,
  opts: {
    bold?: boolean;
    underline?: boolean;
    size?: number;
    italics?: boolean;
    highlight?: 'yellow';
  } = {}
): TextRun {
  return new TextRun({
    text,
    font: FONT,
    size: opts.size ?? SZ_BODY,
    bold: opts.bold,
    italics: opts.italics,
    highlight: opts.highlight,
    underline: opts.underline ? { type: UnderlineType.SINGLE } : undefined,
  });
}

function p(
  children: TextRun[],
  opts: {
    align?: (typeof AlignmentType)[keyof typeof AlignmentType];
    before?: number;
    after?: number;
    indentLeft?: number;
  } = {}
): Paragraph {
  return new Paragraph({
    alignment: opts.align ?? AlignmentType.LEFT,
    spacing: { before: opts.before ?? 30, after: opts.after ?? 30, line: 240 },
    indent: opts.indentLeft ? { left: opts.indentLeft } : undefined,
    children,
  });
}

function cell(
  paragraphs: Array<Paragraph | Table>,
  opts: {
    widthPct?: number;
    rowSpan?: number;
    columnSpan?: number;
    borders?: ITableCellBorders;
    vAlign?: TableVerticalAlign;
    shading?: string;
  } = {}
): TableCell {
  return new TableCell({
    width: opts.widthPct ? { size: opts.widthPct, type: WidthType.PERCENTAGE } : undefined,
    rowSpan: opts.rowSpan,
    columnSpan: opts.columnSpan,
    borders: opts.borders ?? ALL_DOTTED_BORDERS,
    verticalAlign: opts.vAlign ?? VerticalAlign.TOP,
    shading: opts.shading ? { fill: opts.shading } : undefined,
    margins: { top: 45, bottom: 45, left: 75, right: 75 },
    children: paragraphs,
  });
}

function simpleCell(
  text: string,
  opts: {
    bold?: boolean;
    align?: (typeof AlignmentType)[keyof typeof AlignmentType];
    size?: number;
    widthPct?: number;
    rowSpan?: number;
    columnSpan?: number;
    borders?: ITableCellBorders;
    shading?: string;
  } = {}
): TableCell {
  const lines = (text || '').split('\n');
  const paras = lines.map((line) =>
    p([txt(line, { bold: opts.bold, size: opts.size ?? SZ_SMALL })], {
      align: opts.align ?? AlignmentType.LEFT,
      before: 15,
      after: 15,
    })
  );
  return cell(paras, {
    widthPct: opts.widthPct,
    rowSpan: opts.rowSpan,
    columnSpan: opts.columnSpan,
    borders: opts.borders,
    shading: opts.shading,
  });
}

async function fetchImageBytes(url: string): Promise<Uint8Array | null> {
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    const buf = await res.arrayBuffer();
    return new Uint8Array(buf);
  } catch {
    return null;
  }
}

function valOrToken(ph: Record<string, string>, code: string): string {
  const v = ph[code];
  return v && v.trim() !== '' ? v.trim() : code;
}

export async function generateICMemoDocxBlob(memo: ResolvedICMemo): Promise<Blob> {
  const ph = memo.placeholders;
  const bp = memo.branchProposal;

  const [logoBytes, bismillahBytes] = await Promise.all([
    fetchImageBytes('/assets/sib_logo.jpg'),
    fetchImageBytes('/assets/bismillah.jpg'),
  ]);

  const v = (code: string) => valOrToken(ph, code);

  // Header rendered on every page
  const headerChildren: Array<Paragraph | Table> = [];
  if (bismillahBytes) {
    headerChildren.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 40 },
        children: [
          new ImageRun({
            data: bismillahBytes,
            type: 'jpg',
            transformation: { width: 190, height: 22 },
          }),
        ],
      })
    );
  }

  headerChildren.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            cell(
              logoBytes
                ? [
                    new Paragraph({
                      alignment: AlignmentType.LEFT,
                      children: [
                        new ImageRun({
                          data: logoBytes,
                          type: 'jpg',
                          transformation: { width: 58, height: 65 },
                        }),
                      ],
                    }),
                  ]
                : [p([txt('SIB PLC', { bold: true })])],
              { widthPct: 18, borders: ALL_NO_BORDERS }
            ),
            cell(
              [
                p([txt('SAMMILITO ISLAMI BANK PLC.', { bold: true, underline: true, size: SZ_TITLE })], {
                  align: AlignmentType.CENTER,
                  before: 0,
                  after: 20,
                }),
                p([txt('(State-owned Shariah-based Bank)', { bold: true, underline: true, size: SZ_BODY })], {
                  align: AlignmentType.CENTER,
                  before: 0,
                  after: 20,
                }),
                p([txt('UNION BANK UNIT', { bold: true, underline: true, size: SZ_BODY })], {
                  align: AlignmentType.CENTER,
                  before: 0,
                  after: 20,
                }),
                p([txt('Special Asset Management Division Head Office, Dhaka', { underline: true, size: SZ_SMALL })], {
                  align: AlignmentType.CENTER,
                  before: 0,
                  after: 20,
                }),
              ],
              { widthPct: 60, borders: ALL_NO_BORDERS }
            ),
            cell(
              [
                new Table({
                  width: { size: 100, type: WidthType.PERCENTAGE },
                  rows: [
                    new TableRow({
                      children: [
                        simpleCell('Exit with Waiver', {
                          bold: true,
                          size: SZ_BODY,
                          align: AlignmentType.CENTER,
                          borders: ALL_THIN_BORDERS,
                        }),
                      ],
                    }),
                  ],
                }),
              ],
              { widthPct: 22, borders: ALL_NO_BORDERS }
            ),
          ],
        }),
      ],
    })
  );

  const header = new Header({ children: headerChildren });

  const footer = new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [
          txt('Page ', { size: SZ_SMALL }),
          new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: SZ_SMALL }),
          txt(' of ', { size: SZ_SMALL }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: SZ_SMALL }),
        ],
      }),
    ],
  });

  // Document body children
  const body: Array<Paragraph | Table> = [];

  // Top Metadata Table (Dated, Memo no, To, From)
  body.push(
    new Table({
      width: { size: 52, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            simpleCell('Dated', { bold: true, widthPct: 30, borders: ALL_THIN_BORDERS }),
            simpleCell(':', { bold: true, align: AlignmentType.CENTER, widthPct: 8, borders: ALL_THIN_BORDERS }),
            simpleCell(v('#1#'), { bold: true, widthPct: 62, borders: ALL_THIN_BORDERS }),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('Memo no', { bold: true, borders: ALL_THIN_BORDERS }),
            simpleCell(':', { bold: true, align: AlignmentType.CENTER, borders: ALL_THIN_BORDERS }),
            simpleCell('', { borders: ALL_THIN_BORDERS }),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('To', { bold: true, borders: ALL_THIN_BORDERS }),
            simpleCell(':', { bold: true, align: AlignmentType.CENTER, borders: ALL_THIN_BORDERS }),
            simpleCell('Investment Committee', { bold: true, borders: ALL_THIN_BORDERS }),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('From', { bold: true, borders: ALL_THIN_BORDERS }),
            simpleCell(':', { bold: true, align: AlignmentType.CENTER, borders: ALL_THIN_BORDERS }),
            simpleCell('SAMD', { bold: true, borders: ALL_THIN_BORDERS }),
          ],
        }),
      ],
    }),
    p([], { before: 40, after: 40 })
  );

  // A/C and Branch Bar
  body.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            simpleCell(`A/C: ${v('#2#')} (CIS: ${v('#69#')})`, {
              bold: true,
              widthPct: 60,
              borders: ALL_DOTTED_BORDERS,
            }),
            simpleCell(`Branch: ${v('#3#')}`, {
              bold: true,
              widthPct: 40,
              borders: ALL_DOTTED_BORDERS,
            }),
          ],
        }),
        new TableRow({
          children: [
            cell(
              [
                p([
                  txt(
                    `Subject: The customer proposed for full and final settlement of their liability by allowing waiver of Tk.${v('#4#')} along with profit/compensation to be accrued up to the date of full and final settlement as follows:`,
                    { bold: true, size: SZ_BODY }
                  ),
                ]),
              ],
              { columnSpan: 2, borders: ALL_DOTTED_BORDERS }
            ),
          ],
        }),
      ],
    })
  );

  body.push(
    p([txt('[All Figures in this Memo are in Million BDT Unless Stated Otherwise]', { bold: true, size: SZ_SMALL })], {
      align: AlignmentType.RIGHT,
      before: 40,
      after: 40,
    })
  );

  // Main Settlement Table
  const remarksBlock = `1)Settlement Deadline:\n${v('#28#')}\n\n2) Break up\nPrincipal: ${v('#7#')}\nCharges: ${v('#8#')}\nCOF: ${v('#31#')}\nTotal: ${v('#32#')}\n\n3) Security Value:\nMV: ${v('#33#')}\nFSV: ${v('#34#')}`;

  body.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            simpleCell('Mode', { bold: true, rowSpan: 2, widthPct: 12, align: AlignmentType.CENTER }),
            simpleCell('Limit', { bold: true, columnSpan: 2, widthPct: 14, align: AlignmentType.CENTER }),
            simpleCell('Particulars', { bold: true, rowSpan: 2, widthPct: 21, align: AlignmentType.CENTER }),
            simpleCell(`O/S as on ${v('#68#')}`, { bold: true, widthPct: 12, align: AlignmentType.CENTER }),
            simpleCell('Proposed Waiver', { bold: true, widthPct: 12, align: AlignmentType.CENTER }),
            simpleCell('Settlement Amount', { bold: true, widthPct: 12, align: AlignmentType.CENTER }),
            simpleCell('Remarks', { bold: true, rowSpan: 2, widthPct: 17, align: AlignmentType.CENTER }),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('Original', { bold: true, align: AlignmentType.CENTER, widthPct: 8 }),
            simpleCell('RSD', { bold: true, align: AlignmentType.CENTER, widthPct: 6 }),
            simpleCell('01', { bold: true, align: AlignmentType.CENTER }),
            simpleCell('02', { bold: true, align: AlignmentType.CENTER }),
            simpleCell('03 (1-2)', { bold: true, align: AlignmentType.CENTER }),
          ],
        }),
        new TableRow({
          children: [
            simpleCell(v('#5#'), { bold: true, rowSpan: 7 }),
            simpleCell(v('#6#'), { bold: true, rowSpan: 7, align: AlignmentType.CENTER }),
            simpleCell(bp?.limitRSD || '-', { rowSpan: 7, align: AlignmentType.CENTER }),
            simpleCell('Principal'),
            simpleCell(v('#7#'), { bold: true, align: AlignmentType.CENTER }),
            simpleCell('', { align: AlignmentType.CENTER }),
            simpleCell('', { align: AlignmentType.CENTER }),
            simpleCell(remarksBlock, { bold: true, rowSpan: 7 }),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('Charges (Legal & Others)'),
            simpleCell(v('#8#'), { bold: true, align: AlignmentType.CENTER }),
            simpleCell(''),
            simpleCell(''),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('Gestation Rent'),
            simpleCell('-', { align: AlignmentType.CENTER }),
            simpleCell(''),
            simpleCell(''),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('Charged Profit'),
            simpleCell(v('#10#'), { bold: true, align: AlignmentType.CENTER }),
            simpleCell(''),
            simpleCell(''),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('Compensation'),
            simpleCell(v('#11#'), { bold: true, align: AlignmentType.CENTER }),
            simpleCell(''),
            simpleCell(''),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('Uncharged Compensation'),
            simpleCell(v('#12#'), { bold: true, align: AlignmentType.CENTER }),
            simpleCell(''),
            simpleCell(''),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('Uncharged Profit (for T/L)'),
            simpleCell('-', { align: AlignmentType.CENTER }),
            simpleCell('-', { align: AlignmentType.CENTER }),
            simpleCell('-', { align: AlignmentType.CENTER }),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('Grand-Total', { bold: true, columnSpan: 4, align: AlignmentType.RIGHT }),
            simpleCell(v('#25#'), { bold: true, align: AlignmentType.CENTER }),
            simpleCell(v('#4#'), { bold: true, align: AlignmentType.CENTER }),
            simpleCell(v('#27#'), { bold: true, align: AlignmentType.CENTER }),
            simpleCell(''),
          ],
        }),
      ],
    })
  );

  // Note line
  body.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            simpleCell(
              `Note: Information Regarding Down payment (deposited Tk.${v('#35#')} Million in A/C: ${v('#36#')} on ${v('#68#')}.)`,
              { bold: true, size: SZ_SMALL }
            ),
          ],
        }),
      ],
    })
  );

  // Key Points
  body.push(
    p([txt('Key Points:', { bold: true, underline: true })], { before: 70, after: 20 }),
    p([
      txt('This Memo is placed in conformity with ', { size: SZ_SMALL }),
      txt('BRPD 1 Circular Letter No. 23 dated 29.06.2026', { bold: true, size: SZ_SMALL }),
      txt(' under the following contexts:', { size: SZ_SMALL }),
    ]),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            simpleCell('Particulars', { bold: true, columnSpan: 3, widthPct: 58 }),
            simpleCell('Waiver recommended by:', { bold: true, columnSpan: 3, widthPct: 42 }),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('a.', { widthPct: 5 }),
            simpleCell('Total [Profit+ Compensation+ Uncharged Comp.& Profit]', { widthPct: 40 }),
            simpleCell(`: ${v('#42#')}`, { bold: true, widthPct: 13 }),
            simpleCell('d.', { widthPct: 5 }),
            simpleCell('Client', { widthPct: 22 }),
            simpleCell(`: ${v('#45#')}`, { bold: true, widthPct: 15 }),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('b.'),
            simpleCell('Cost of Fund'),
            simpleCell(`: ${v('#31#')}`, { bold: true }),
            simpleCell('e.'),
            simpleCell('Branch'),
            simpleCell(`: ${v('#45#')}`, { bold: true }),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('c.', { rowSpan: 4 }),
            simpleCell('Maximum Waiver Considering COF [a-b]', { rowSpan: 4 }),
            simpleCell(`: ${v('#44#')}`, { bold: true, rowSpan: 4 }),
            simpleCell('f.'),
            simpleCell('SAMD, HO'),
            simpleCell(`: ${v('#47#')}`, { bold: true }),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('g.'),
            simpleCell('HOIC'),
            simpleCell(`: ${v('#47#')}`, { bold: true }),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('h.'),
            simpleCell('MEC'),
            simpleCell(`: ${v('#47#')}`, { bold: true }),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('i.'),
            simpleCell('IC & CD'),
            simpleCell(`: ${v('#47#')}`, { bold: true }),
          ],
        }),
      ],
    })
  );

  // Justification of The Proposed Facility
  const justLines = memo.justificationLines
    ? memo.justificationLines.split('\n').map((l) => l.trim()).filter(Boolean)
    : [];
  body.push(
    p([txt('Justification of The Proposed Facility:', { bold: true, underline: true })], { before: 70, after: 20 })
  );
  if (justLines.length > 0) {
    body.push(p([txt('a. ', { bold: true }), txt(justLines.join(' '))], { indentLeft: 180 }));
  } else {
    body.push(p([txt(`a. ${v('#51#')}`)], { indentLeft: 180 }));
  }
  body.push(
    p([txt('b. Investment was not created against any fraud and forgery.')], { indentLeft: 180 }),
    p([txt('c. The client is not a willful defaulter.')], { indentLeft: 180 }),
    p([txt('Previous Board Decision(s) Regarding Rescheduling/Waiver/Written-Off: Nil', { bold: true })], {
      before: 50,
      after: 40,
    })
  );

  // Section 1
  body.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            simpleCell(`1. Nature of Business: ${v('#54#')}`, { bold: true, size: SZ_BODY }),
          ],
        }),
      ],
    })
  );

  // Section 2
  body.push(
    p([txt(`2. Liability Position of the Clients and Its Sister Concerns with Our Bank as on ${v('#68#')}:`, { bold: true })], {
      before: 60,
      after: 20,
    }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            simpleCell('Name of the Concern', { bold: true, widthPct: 14 }),
            simpleCell('Mode of Investment', { bold: true, widthPct: 16 }),
            simpleCell('Limit', { bold: true, widthPct: 10, align: AlignmentType.CENTER }),
            simpleCell('Outstanding', { bold: true, widthPct: 12, align: AlignmentType.CENTER }),
            simpleCell('Validity', { bold: true, widthPct: 12, align: AlignmentType.CENTER }),
            simpleCell('Overdue', { bold: true, widthPct: 10, align: AlignmentType.CENTER }),
            simpleCell('Recovery in existing deals till date [Mode wise]', { bold: true, widthPct: 16, align: AlignmentType.CENTER }),
            simpleCell('Remarks', { bold: true, widthPct: 10 }),
          ],
        }),
        new TableRow({
          children: [
            simpleCell(v('#2#'), { bold: true, rowSpan: 2 }),
            simpleCell(`${v('#55#')}\nFunded`, { bold: true }),
            simpleCell(v('#56#'), { bold: true, align: AlignmentType.CENTER }),
            simpleCell(v('#25#'), { bold: true, align: AlignmentType.CENTER }),
            simpleCell(v('#58#'), { bold: true, align: AlignmentType.CENTER }),
            simpleCell(v('#59#'), { bold: true, align: AlignmentType.CENTER }),
            simpleCell(v('#60#'), { bold: true, align: AlignmentType.CENTER }),
            simpleCell(bp?.sec2Remarks || ''),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('Non Funded', { bold: true }),
            simpleCell(''),
            simpleCell(''),
            simpleCell(''),
            simpleCell(''),
            simpleCell(''),
            simpleCell(''),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('Total', { bold: true, columnSpan: 2, align: AlignmentType.RIGHT }),
            simpleCell(v('#56#'), { bold: true, align: AlignmentType.CENTER }),
            simpleCell(v('#25#'), { bold: true, align: AlignmentType.CENTER }),
            simpleCell(v('#58#'), { bold: true, align: AlignmentType.CENTER }),
            simpleCell(v('#59#'), { bold: true, align: AlignmentType.CENTER }),
            simpleCell(v('#60#'), { bold: true, align: AlignmentType.CENTER }),
            simpleCell(''),
          ],
        }),
      ],
    })
  );

  // Section 3 (from Branch Proposal)
  const sec3Rows = bp?.section3Rows && bp.section3Rows.length > 0
    ? bp.section3Rows
    : [
        { sl: 'a.', nameOfConcern: '', bankOrNbfi: '', natureOfFacility: '', limit: '', outstanding: '', overdueEol: '', expiryDate: '', clStatus: '' },
        { sl: 'b.', nameOfConcern: '', bankOrNbfi: '', natureOfFacility: '', limit: '', outstanding: '', overdueEol: '', expiryDate: '', clStatus: '' },
        { sl: 'c.', nameOfConcern: '', bankOrNbfi: '', natureOfFacility: '', limit: '', outstanding: '', overdueEol: '', expiryDate: '', clStatus: '' },
      ];
  body.push(
    p([txt('3. Liability of the Client & Its Sister/Allied Concerns with Other Units of SIB:', { bold: true, underline: true })], {
      before: 70,
      after: 20,
    }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            simpleCell('SL', { bold: true, widthPct: 5 }),
            simpleCell('Name of Concern', { bold: true, widthPct: 14 }),
            simpleCell('Name of the Banks/NBFIs', { bold: true, widthPct: 17 }),
            simpleCell('Nature of Facility', { bold: true, widthPct: 12 }),
            simpleCell('Limit', { bold: true, widthPct: 9 }),
            simpleCell('Outstanding with profit/Interest', { bold: true, widthPct: 15 }),
            simpleCell('Overdue/EOL', { bold: true, widthPct: 10 }),
            simpleCell('Expiry Date', { bold: true, widthPct: 10 }),
            simpleCell('CL Status', { bold: true, widthPct: 8 }),
          ],
        }),
        ...sec3Rows.map(
          (r) =>
            new TableRow({
              children: [
                simpleCell(r.sl),
                simpleCell(r.nameOfConcern),
                simpleCell(r.bankOrNbfi),
                simpleCell(r.natureOfFacility),
                simpleCell(r.limit),
                simpleCell(r.outstanding),
                simpleCell(r.overdueEol),
                simpleCell(r.expiryDate),
                simpleCell(r.clStatus),
              ],
            })
        ),
      ],
    })
  );

  // Section 4 (CIB Status: Report dated #62# shows, table from Branch Proposal)
  const defaultCibStatuses = ['STD', 'SMA', 'SS', 'DF', 'BL', 'Willful Default (WD)', 'Write-Off', 'Non-Funded', 'Stay Order', 'TOTAL'];
  const sec4Rows = bp?.section4Rows && bp.section4Rows.length > 0
    ? bp.section4Rows
    : defaultCibStatuses.map((s) => ({
        status: s,
        unitBorrower: '',
        unitGuarantor: '',
        otherBorrower: '',
        otherGuarantor: '',
        remarks: '',
      }));

  body.push(
    p([
      txt('4. CIB Status: ', { bold: true, underline: true }),
      txt(`Report dated ${v('#62#')} shows:`, { bold: true }),
    ], { before: 70, after: 20 }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            simpleCell('STATUS', { bold: true, rowSpan: 2, widthPct: 18 }),
            simpleCell('With respective unit', { bold: true, columnSpan: 2, widthPct: 26, align: AlignmentType.CENTER }),
            simpleCell('With Other Bank', { bold: true, columnSpan: 2, widthPct: 26, align: AlignmentType.CENTER }),
            simpleCell('Remarks\n(Pls. give concern wise Breakup)', { bold: true, rowSpan: 2, widthPct: 30, align: AlignmentType.CENTER }),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('Borrower', { bold: true, align: AlignmentType.CENTER }),
            simpleCell('Guarantor', { bold: true, align: AlignmentType.CENTER }),
            simpleCell('Borrower', { bold: true, align: AlignmentType.CENTER }),
            simpleCell('Guarantor', { bold: true, align: AlignmentType.CENTER }),
          ],
        }),
        ...sec4Rows.map(
          (r) =>
            new TableRow({
              children: [
                simpleCell(r.status, { bold: /TOTAL/i.test(r.status) }),
                simpleCell(r.unitBorrower, { align: AlignmentType.CENTER, bold: /TOTAL/i.test(r.status) }),
                simpleCell(r.unitGuarantor, { align: AlignmentType.CENTER }),
                simpleCell(r.otherBorrower, { align: AlignmentType.CENTER }),
                simpleCell(r.otherGuarantor, { align: AlignmentType.CENTER }),
                simpleCell(r.remarks),
              ],
            })
        ),
      ],
    })
  );

  // Section 5 (Legal Action Status)
  const sec5Rows = bp?.section5Rows && bp.section5Rows.length > 0
    ? bp.section5Rows
    : [{ sl: '', typeOfCase: '', caseNo: '', filingDate: '', amount: '', previousDate: '', nextDate: '', presentStatus: '' }];
  body.push(
    p([
      txt('5. Legal Action Status as on Memo Date:', { bold: true, underline: true }),
      ...(bp?.section5StatusNote ? [txt(` ${bp.section5StatusNote}`, { bold: true })] : []),
    ], { before: 70, after: 20 }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            simpleCell('Sl', { bold: true, widthPct: 5 }),
            simpleCell('Type of Case', { bold: true, widthPct: 16 }),
            simpleCell('Case No.', { bold: true, widthPct: 13 }),
            simpleCell('Filling date', { bold: true, widthPct: 12 }),
            simpleCell('Amount', { bold: true, widthPct: 12 }),
            simpleCell('Previous date', { bold: true, widthPct: 12 }),
            simpleCell('Next date', { bold: true, widthPct: 12 }),
            simpleCell('Present status', { bold: true, widthPct: 18 }),
          ],
        }),
        ...sec5Rows.map(
          (r) =>
            new TableRow({
              children: [
                simpleCell(r.sl),
                simpleCell(r.typeOfCase),
                simpleCell(r.caseNo),
                simpleCell(r.filingDate),
                simpleCell(r.amount),
                simpleCell(r.previousDate),
                simpleCell(r.nextDate),
                simpleCell(r.presentStatus),
              ],
            })
        ),
      ],
    })
  );

  // Section 6 (Audit, Compliance)
  const defaultSec6 = [
    { sl: 'a.', authority: 'IC & CD', lastInspectionDate: '', objectionStatus: '', objections: '', branchReply: '' },
    { sl: 'b.', authority: 'DBI, Bangladesh Bank', lastInspectionDate: '', objectionStatus: 'No', objections: '', branchReply: '' },
    { sl: 'c.', authority: 'DOS, Bangladesh Bank', lastInspectionDate: 'Not done', objectionStatus: 'No', objections: '', branchReply: '' },
    { sl: 'd.', authority: 'BFIU, Bangladesh Bank', lastInspectionDate: 'Not done', objectionStatus: 'No', objections: '', branchReply: '' },
    { sl: 'e', authority: 'BRD, Bangladesh Bank', lastInspectionDate: 'Not done', objectionStatus: 'No', objections: '', branchReply: '' },
    { sl: 'f', authority: 'ACC(Anti-Corruption Commission)', lastInspectionDate: 'Not done', objectionStatus: 'No', objections: '', branchReply: '' },
  ];
  const sec6Rows = bp?.section6Rows && bp.section6Rows.length > 0 ? bp.section6Rows : defaultSec6;
  body.push(
    p([
      txt('6. Audit, Compliance and Regulatory Authority Related Status:', { bold: true, underline: true }),
      ...(bp?.section6StatusNote ? [txt(` ${bp.section6StatusNote}`, { bold: true })] : []),
    ], { before: 70, after: 20 }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            simpleCell('SL', { bold: true, widthPct: 5 }),
            simpleCell('Name of the Authority', { bold: true, widthPct: 24 }),
            simpleCell('Last Inspection / Report Date', { bold: true, widthPct: 18 }),
            simpleCell('Objection Status (Yes/No)', { bold: true, widthPct: 15 }),
            simpleCell('Objections (If yes)', { bold: true, widthPct: 19 }),
            simpleCell('Branch Reply/ Current Position', { bold: true, widthPct: 19 }),
          ],
        }),
        ...sec6Rows.map(
          (r) =>
            new TableRow({
              children: [
                simpleCell(r.sl),
                simpleCell(r.authority),
                simpleCell(r.lastInspectionDate),
                simpleCell(r.objectionStatus),
                simpleCell(r.objections),
                simpleCell(r.branchReply),
              ],
            })
        ),
      ],
    })
  );

  // Section 7 (Security)
  const defaultSec7Primary = [
    { particulars: 'i) Hypothecation of Stock as on ………………..', value: '' },
    { particulars: 'ii) Hypothecation of Machinery', value: '' },
    { particulars: 'iii) Assignment of Bills against Work Order', value: '' },
    { particulars: 'iv) Margin held against existing LC & BG liability as on……..', value: '' },
  ];
  const sec7Primary = bp?.section7PrimaryRows && bp.section7PrimaryRows.length > 0 ? bp.section7PrimaryRows : defaultSec7Primary;
  const defaultSchedules = [
    'Schedule-01: Full Schedule with owner name as per Mortgage Deed.',
    'Schedule-02: Full Schedule with owner name as per Mortgage Deed.',
    'Schedule-03: Full Schedule with owner name as per Mortgage Deed.',
    'Schedule-04: Full Schedule with owner name as per Mortgage Deed.',
  ];
  const schedules = bp?.section7CollateralSchedules && bp.section7CollateralSchedules.length > 0 ? bp.section7CollateralSchedules : defaultSchedules;
  const defaultValRows = ['Schedule-01', 'Schedule-02', 'Schedule-03', 'Schedule-04', 'Building', 'Other', 'Grand Total'].map(
    (s) => ({
      schedule: s,
      origBranchMv: '',
      origBranchFsv: '',
      origSurvMv: '',
      origSurvFsv: '',
      presBranchMv: '',
      presBranchFsv: '',
      presSurvMv: '',
      presSurvFsv: '',
      presAvMv: '',
      presAvFsv: '',
    })
  );
  const valRows = bp?.section7ValuationRows && bp.section7ValuationRows.length > 0 ? bp.section7ValuationRows : defaultValRows;

  body.push(
    p([txt('7. Security:', { bold: true, underline: true })], { before: 70, after: 15 }),
    p([
      txt('a. Primary:', { bold: true, underline: true }),
      ...(bp?.section7PrimaryNote ? [txt(` ${bp.section7PrimaryNote}`, { bold: true })] : []),
    ], { indentLeft: 120 }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            simpleCell('Particulars', { bold: true, widthPct: 75 }),
            simpleCell('Value', { bold: true, widthPct: 25, align: AlignmentType.CENTER }),
          ],
        }),
        ...sec7Primary.map(
          (r) =>
            new TableRow({
              children: [simpleCell(r.particulars), simpleCell(r.value, { align: AlignmentType.CENTER })],
            })
        ),
        new TableRow({
          children: [
            simpleCell('Total value of primary security', { bold: true, align: AlignmentType.CENTER }),
            simpleCell(bp?.section7PrimaryTotal || '', { bold: true, align: AlignmentType.CENTER }),
          ],
        }),
      ],
    }),
    p([
      txt('b. Collateral:', { bold: true, underline: true }),
      ...(bp?.section7CollateralNote ? [txt(` ${bp.section7CollateralNote}`, { bold: true })] : []),
    ], { indentLeft: 120, before: 40 }),
    ...schedules.map((s) => p([txt(s, { size: SZ_SMALL })], { indentLeft: 120 })),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            simpleCell('Schedule', { bold: true, rowSpan: 3, widthPct: 14 }),
            simpleCell('Value of the mortgaged properties as assessed', { bold: true, columnSpan: 10, align: AlignmentType.CENTER }),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('At the time of allowing investment', { bold: true, columnSpan: 4, align: AlignmentType.CENTER }),
            simpleCell('At present', { bold: true, columnSpan: 6, align: AlignmentType.CENTER }),
          ],
        }),
        new TableRow({
          children: [
            simpleCell('By Branch MV', { bold: true, size: 14 }),
            simpleCell('By Branch FSV', { bold: true, size: 14 }),
            simpleCell('Surveyor MV', { bold: true, size: 14 }),
            simpleCell('Surveyor FSV', { bold: true, size: 14 }),
            simpleCell('By Branch MV', { bold: true, size: 14 }),
            simpleCell('By Branch FSV', { bold: true, size: 14 }),
            simpleCell('Surveyor MV', { bold: true, size: 14 }),
            simpleCell('Surveyor FSV', { bold: true, size: 14 }),
            simpleCell('By AV Cell MV', { bold: true, size: 14 }),
            simpleCell('By AV Cell FSV', { bold: true, size: 14 }),
          ],
        }),
        ...valRows.map(
          (r) =>
            new TableRow({
              children: [
                simpleCell(r.schedule, { bold: /Grand/i.test(r.schedule) }),
                simpleCell(r.origBranchMv),
                simpleCell(r.origBranchFsv),
                simpleCell(r.origSurvMv),
                simpleCell(r.origSurvFsv),
                simpleCell(r.presBranchMv),
                simpleCell(r.presBranchFsv),
                simpleCell(r.presSurvMv),
                simpleCell(r.presSurvFsv),
                simpleCell(r.presAvMv),
                simpleCell(r.presAvFsv),
              ],
            })
        ),
        new TableRow({
          children: [
            simpleCell(`Outstanding to Lowest Security value Coverage (%): ${bp?.section7CoveragePct || '……………%'}`, {
              bold: true,
              columnSpan: 11,
              align: AlignmentType.CENTER,
            }),
          ],
        }),
      ],
    }),
    p([txt('c. Others:', { bold: true })], { indentLeft: 120, before: 40 }),
    p([txt('a. Usual Charge documents.', { bold: true })], { indentLeft: 260 }),
    p([txt('b. Acceptance of Branch Sanction by the Client.', { bold: true })], { indentLeft: 260 })
  );

  // Section 8 (Source of Repayment from Manual Inputs)
  body.push(
    p([
      txt('8. Source of Repayment: ', { bold: true, underline: true }),
      txt(memo.sourceOfRepayment || ''),
    ], { before: 60, after: 40 })
  );

  // Section 9 (Sequence of Head Office Sanction(S))
  const sec9Rows = bp?.section9Rows && bp.section9Rows.length > 0
    ? bp.section9Rows
    : ['a.', 'b.', 'c.', 'd.', 'e.', 'f.', 'g.'].map((sl) => ({
        sl,
        referenceNo: '',
        date: '',
        limit: '',
        remarks: '',
      }));
  body.push(
    p([txt('9. Sequence of Head Office Sanction(S):', { bold: true, underline: true })], { before: 60, after: 20 }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            simpleCell('SL', { bold: true, widthPct: 8 }),
            simpleCell('HO Sanction Reference No.', { bold: true, widthPct: 36 }),
            simpleCell('Date', { bold: true, widthPct: 16 }),
            simpleCell('Limit', { bold: true, widthPct: 14 }),
            simpleCell('Remarks', { bold: true, widthPct: 26 }),
          ],
        }),
        ...sec9Rows.map(
          (r) =>
            new TableRow({
              children: [
                simpleCell(r.sl),
                simpleCell(r.referenceNo),
                simpleCell(r.date),
                simpleCell(r.limit),
                simpleCell(r.remarks),
              ],
            })
        ),
      ],
    })
  );

  // Section 10 (Brief Particulars of The Client)
  const sp = bp?.section10Particulars;
  const sec10Items: Array<[string, string, string]> = [
    ['a.', 'Name of the Client', sp?.nameOfClient || ''],
    ['b.', 'Unique Customer ID', sp?.uniqueCustomerId || ''],
    ['c.', 'AWCD/MSD Account with opening date', sp?.awcdMsdAccount || ''],
    ['d.', 'Business Type/Nature', sp?.businessTypeNature || ''],
    ['e.', 'Business Management Type', sp?.businessManagementType || ''],
    ['f.', 'Client Type', sp?.clientType || ''],
    ['g.', 'Key Person with Phone No.', sp?.keyPersonPhone || ''],
    ['', 'Address', sp?.address || ''],
    ['h.', 'Office/Business Address', sp?.officeBusinessAddress || ''],
    ['', 'Factory Address', sp?.factoryAddress || ''],
    ['i.', 'Purpose of Investment', sp?.purposeOfInvestment || ''],
    ['j.', 'Limit of Investment', sp?.limitOfInvestment || ''],
    ['k.', 'Mode of Investment', sp?.modeOfInvestment || ''],
    ['l.', 'Last Visit Date', sp?.lastVisitDate || ''],
    ['m.', 'Stock Position', sp?.stockPosition || ''],
    ['n.', 'Fixed Asset', sp?.fixedAsset || ''],
    ['o.', 'Business Net Worth', sp?.businessNetWorth || ''],
    ['p.', 'Insurance Coverage', sp?.insuranceCoverage || ''],
    ['q.', 'Trade License & Validity', sp?.tradeLicenseValidity || ''],
    ['r.', 'Other Regulatory Documents', sp?.otherRegulatoryDocuments || ''],
    ['s.', 'Date of Original Investment', sp?.dateOfOriginalInvestment || ''],
    ['t.', '1st Disbursement Manager', sp?.firstDisbursementManager || ''],
    ['u.', 'First Date of Classification with Status', sp?.firstDateOfClassification || ''],
  ];

  body.push(
    p([txt('10. Brief Particulars of The Client:', { bold: true, underline: true })], { before: 70, after: 20 }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: sec10Items.map(
        ([idx, label, val]) =>
          new TableRow({
            children: [
              simpleCell(idx, { widthPct: 6 }),
              simpleCell(label, { widthPct: 36 }),
              simpleCell(':', { widthPct: 4, align: AlignmentType.CENTER }),
              simpleCell(val, { widthPct: 54 }),
            ],
          })
      ),
    })
  );

  // Particulars of Proprietor/Partner/Directors
  const dirRows = bp?.section10Directors && bp.section10Directors.length > 0
    ? bp.section10Directors
    : [{ name: '', age: '', education: '', designation: '', sharePct: '', experience: '', netWorth: '' }];
  body.push(
    p([txt('•  Particular of the Proprietor/Partner/Directors:', { bold: true, underline: true })], { before: 50, after: 20 }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            simpleCell('Name of the Directors', { bold: true, widthPct: 24 }),
            simpleCell('Age', { bold: true, widthPct: 8 }),
            simpleCell('Education', { bold: true, widthPct: 14 }),
            simpleCell('Designation', { bold: true, widthPct: 16 }),
            simpleCell('Share (%)', { bold: true, widthPct: 12 }),
            simpleCell('Experience', { bold: true, widthPct: 14 }),
            simpleCell('Net worth', { bold: true, widthPct: 12 }),
          ],
        }),
        ...dirRows.map(
          (d) =>
            new TableRow({
              children: [
                simpleCell(d.name),
                simpleCell(d.age),
                simpleCell(d.education),
                simpleCell(d.designation),
                simpleCell(d.sharePct),
                simpleCell(d.experience),
                simpleCell(d.netWorth),
              ],
            })
        ),
      ],
    })
  );

  // Remarks: Director of any Bank/FI/Insurance
  const bankDirRows = bp?.section10BankDirectors && bp.section10BankDirectors.length > 0
    ? bp.section10BankDirectors
    : [
        { name: '', bankNbfi: '', designation: '', sharePct: '', numShares: '' },
        { name: '', bankNbfi: '', designation: '', sharePct: '', numShares: '' },
      ];
  body.push(
    p([txt('•  Remarks: Director of any Bank/ FI/ Insurance company:', { bold: true, underline: true })], {
      before: 50,
      after: 20,
    }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            simpleCell('Name of the Director', { bold: true, widthPct: 24 }),
            simpleCell('Name of the bank/NBFIs', { bold: true, widthPct: 24 }),
            simpleCell('Designation/Position', { bold: true, widthPct: 18 }),
            simpleCell('Percentage of Shares', { bold: true, widthPct: 17 }),
            simpleCell('Number of Shares', { bold: true, widthPct: 17 }),
          ],
        }),
        ...bankDirRows.map(
          (d) =>
            new TableRow({
              children: [
                simpleCell(d.name),
                simpleCell(d.bankNbfi),
                simpleCell(d.designation),
                simpleCell(d.sharePct),
                simpleCell(d.numShares),
              ],
            })
        ),
      ],
    })
  );

  // Section 11 (Financials of The Client)
  const finRows = bp?.section11Rows && bp.section11Rows.length > 0
    ? bp.section11Rows
    : [
        { year: '', grossSales: bp?.section11FinancialsNote || '', grossProfit: '', netProfit: '', totalAsset: '', totalLiabilities: '', netWorth: '' },
      ];
  body.push(
    p([txt('11. Financials of The Client:', { bold: true, underline: true })], { before: 70, after: 20 }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            simpleCell('Year', { bold: true, widthPct: 10 }),
            simpleCell('Gross Sales', { bold: true, widthPct: 15 }),
            simpleCell('Gross Profit', { bold: true, widthPct: 15 }),
            simpleCell('Net Profit', { bold: true, widthPct: 15 }),
            simpleCell('Total Asset', { bold: true, widthPct: 15 }),
            simpleCell('Total liabilities', { bold: true, widthPct: 15 }),
            simpleCell('Net Worth', { bold: true, widthPct: 15 }),
          ],
        }),
        ...finRows.map(
          (r) =>
            new TableRow({
              children: [
                simpleCell(r.year),
                simpleCell(r.grossSales),
                simpleCell(r.grossProfit),
                simpleCell(r.netProfit),
                simpleCell(r.totalAsset),
                simpleCell(r.totalLiabilities),
                simpleCell(r.netWorth),
              ],
            })
        ),
      ],
    })
  );

  // Section 12 (Branch Visit Report)
  body.push(
    p([
      txt('12. Branch Visit Report: ', { bold: true, underline: true }),
      txt(memo.branchVisitReport || ''),
    ], { before: 70, after: 30 })
  );

  // Section 13 (Recommendation of The Branch)
  body.push(
    p([
      txt('13. Recommendation of The Branch: ', { bold: true, underline: true }),
      txt(`The Branch Recommended the following vide Proposal Ref: ${v('#64#')} dated ${v('#65#')}`),
    ], { before: 50, after: 30 })
  );

  // Section 14
  body.push(
    p([
      txt('14. Recommendation of Internal Control & Compliance Division: ', { bold: true, underline: true }),
      txt('N/A', { bold: true }),
    ], { before: 50, after: 30 })
  );

  // Section 15 (Observation and Recommendation of SAMD)
  body.push(
    p([txt('15. Observation and Recommendation of SAMD:', { bold: true, underline: true })], { before: 50, after: 20 })
  );
  if (memo.samdObservation) {
    for (const line of memo.samdObservation.split('\n').filter((l) => l.trim())) {
      body.push(p([txt(line.trim())], { indentLeft: 180 }));
    }
  }

  // Section 16 (Particulars of the facility)
  const sec16Rows: Array<[string, string, string, boolean]> = [
    ['a.', 'Facility', 'EXIT with Waiver under BRPD 1 Circular Letter No. 23 dated 29.06.2026', true],
    ['b.', 'Outstanding', `Tk. ${v('#25#')} million as on ${v('#68#')}.`, true],
    ['c.', 'Down Payment', `Down-payment of Tk. ${v('#35#')} Million is kept in ${v('#36#')} A/C ${v('#36#')}`, true],
    ['d.', 'Waiver', `Tk. ${v('#4#')} million as on ${v('#68#')}`, true],
    ['e.', 'Settlement/Exit amount', `Tk. ${v('#27#')} Million as on ${v('#68#')}`, true],
    ['f.', 'Settlement period', `${v('#28#')}`, true],
    ['g.', 'Profit/Rent rate', 'Without charging further profit.', false],
    ['h.', 'Repayment Method', 'Lump sum.', false],
  ];
  body.push(
    p([txt('16. Particulars of the facility:', { bold: true, underline: true })], { before: 60, after: 20 }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: sec16Rows.map(
        ([sl, label, val, isBold]) =>
          new TableRow({
            children: [
              simpleCell(sl, { widthPct: 5 }),
              simpleCell(label, { widthPct: 24 }),
              simpleCell(val, { bold: isBold, widthPct: 71 }),
            ],
          })
      ),
    })
  );

  // Section 17 (Special Term(s) & Condition(s))
  const terms: Array<[string, string, boolean]> = [
    ['i.', 'The waiver of profit shall be adjusted from FAD, Head Office from the provision amount kept (if any) and the rest amount shall be adjusted by branch gradually after adjustment of the settlement amount within the approved time frame.', false],
    ['ii.', 'The final settlement amount will be determined at the date of execution of settlement considering up to date liability position.', false],
    ['iii.', 'The Branch must ensure adjustment of all other direct and indirect liabilities of the client before allowing the final redemption of the mortgaged collateral security(s).', false],
    ['iv.', 'The Branch must ensure adjustment of all non-funded liabilities (L/C, Bank Guarantee, Credit Card etc.) before allowing final redemption of the mortgaged collateral security(s). The Branch must also ensure that any encashable security (Margin for Bank Guarantee, L/C, MTDR, Scheme deposit in the name of the client) including collateral is properly mentioned in the proposal. If any encashable security (Margin for Bank Guarantee, L/C, MTDR, Scheme deposit in the name of the client) is found not mentioned in the proposal and later encashed in favor of the client, the Manager will be held responsible.', false],
    ['v.', 'If the client fails to repay or fails to adjust the liability within the given time or fails to pay installments regularly, all facilities shall be cancelled and legal action shall be initiated by the Branch.', false],
    ['vi.', 'Branch should be utmost sincere to protect bank interest in the course of execution partial redemption.', false],
    ['vii.', 'Branch is advised to follow the guidelines of Bangladesh Bank BRPD circular No. 06/22, Circular Letter No.18/22, 46/22 and subsequent up-dates as well as all instruction circular(s) of SIBL.', false],
    ['viii.', 'The Profit rate of the stipulated investment shall not be changed or reduced to 0.00% until full and final adjustment of the Investment, as approved by Head Office.', true],
    ['ix.', 'All the legal expenses and other charges (if any) already incurred and to be incurred in connection with the withdrawal of case/suits and others issues shall be borne by the client.', false],
  ];

  body.push(
    p([txt('17. Special Term(s) & Condition(s):', { bold: true, underline: true })], { before: 70, after: 20 }),
    p([txt('If the proposal approved, the branch is advised to comply the following special terms and conditions:')], {
      indentLeft: 180,
    }),
    ...terms.map(([num, clause, isBold]) =>
      p([txt(`${num}  `, { bold: isBold }), txt(clause, { bold: isBold })], {
        indentLeft: 360,
        before: 25,
        after: 25,
      })
    )
  );

  // Closing box and Signatories
  body.push(
    p([], { before: 60, after: 60 }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            simpleCell('The proposal is hereby placed before the Investment Committee for kind recommendation.', {
              bold: true,
              align: AlignmentType.CENTER,
              size: SZ_BODY,
              borders: ALL_THIN_BORDERS,
            }),
          ],
        }),
      ],
    }),
    p([txt('Union Bank PLC Unit', { bold: true })], { before: 70, after: 10 }),
    p([txt('Special Asset Management Division', { bold: true })], { before: 0, after: 60 }),
    p([txt('Submitted for kind decision.', { bold: true })], { align: AlignmentType.CENTER, before: 40, after: 120 }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            cell(
              [
                p([txt('(Shahidur Rahman khan)', { bold: true })], { before: 200, after: 10 }),
                p([txt('Deputy Managing Director & CRO')], { before: 0, after: 10 }),
                p([txt('Sammilito Islami Bank PLC.')], { before: 0, after: 10 }),
              ],
              { widthPct: 50, borders: ALL_NO_BORDERS }
            ),
            cell(
              [
                p([txt('(Ahmed Zafrul Hasan)', { bold: true })], { before: 200, after: 10 }),
                p([txt('Additional Managing Director & COO')], { before: 0, after: 10 }),
                p([txt('Sammilito Islami Bank PLC.')], { before: 0, after: 10 }),
              ],
              { widthPct: 50, borders: ALL_NO_BORDERS }
            ),
          ],
        }),
      ],
    }),
    p([txt('(Abedur Rahman Sikder)', { bold: true })], { align: AlignmentType.CENTER, before: 260, after: 10 }),
    p([txt('Managing Director')], { align: AlignmentType.CENTER, before: 0, after: 10 }),
    p([txt('Sammilito Islami Bank PLC.')], { align: AlignmentType.CENTER, before: 0, after: 10 }),
    p([txt('(Transferee of the assets and liabilities of')], { align: AlignmentType.CENTER, before: 0, after: 10 }),
    p([txt('EXIM Bank, SIBL, FSIB, Global Islami')], { align: AlignmentType.CENTER, before: 0, after: 10 }),
    p([txt('Bank and Union Bank)')], { align: AlignmentType.CENTER, before: 0, after: 10 })
  );

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1500,
              bottom: 720,
              left: 820,
              right: 820,
              header: 360,
              footer: 360,
            },
          },
        },
        headers: { default: header },
        footers: { default: footer },
        children: body,
      },
    ],
  });

  return await Packer.toBlob(doc);
}
