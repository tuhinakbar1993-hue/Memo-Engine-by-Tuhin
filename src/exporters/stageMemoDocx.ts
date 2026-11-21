import {
  AlignmentType,
  Document,
  Packer,
  Paragraph,
  Table,
  TableLayoutType,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';
import {
  buildBoardMecRecommendationText,
  buildMecIcRecommendationText,
  formatChronologicOrderWithYear,
} from '../engines/icMemoEngine';
import type { ResolvedStageMemo } from '../types/memo';
import {
  A4_HEIGHT_TWIPS,
  A4_WIDTH_TWIPS,
  ALL_THIN_BORDERS,
  BODY_WIDTH_TWIPS,
  MARGIN_LR_TWIPS,
  SZ_BODY,
  SZ_META,
  buildBankHeaderAndFooter,
  buildICMemoCoreBlocks,
  cell,
  p,
  pctWidths,
  simpleCell,
  txt,
  valOrToken,
} from './icMemoDocx';

function renderNarrativeParagraphs(
  heading: string,
  text: string,
  opts: { before?: number; after?: number; fallback?: string } = {}
): Paragraph[] {
  const blocks: Paragraph[] = [
    p([txt(heading, { bold: true, underline: true })], {
      before: opts.before ?? 45,
      after: 15,
    }),
  ];
  const content = (text || '').trim() || opts.fallback || '';
  if (content) {
    for (const line of content.split('\n').filter((l) => l.trim())) {
      blocks.push(
        p([txt(line.trim())], {
          indentLeft: 180,
          before: 12,
          after: opts.after ?? 16,
          line: 245,
        })
      );
    }
  }
  return blocks;
}

/**
 * Renders the prescribed Committee Recommendation heading + paragraph, bolding the
 * `[Order]/26` and `[Date]` segments in "in its <ORDER> meeting held on <DATE>."
 */
function renderCommitteeRecommendationBlocks(
  heading: string,
  text: string,
  opts: { before?: number; after?: number } = {}
): Paragraph[] {
  const blocks: Paragraph[] = [
    p([txt(heading, { bold: true, underline: true })], {
      before: opts.before ?? 45,
      after: 15,
    }),
  ];

  const lines = (text || '')
    .split('\n')
    .map((l) => l.trim())
    .filter(Boolean);

  for (const line of lines) {
    const match = line.match(
      /^(.*?\bin\s+its\s+)(.+?)(\s+meeting\s+held\s+on\s+)(.+?)(\.\s*The\s+Committee\s+recommended\s+the\s+proposal\s+as\s+presented\.?.*)$/i
    );
    if (match) {
      const runs: TextRun[] = [
        txt(match[1]),
        txt(match[2], { bold: true }),
        txt(match[3]),
        txt(match[4], { bold: true }),
        txt(match[5]),
      ];
      blocks.push(
        p(runs, {
          indentLeft: 180,
          before: 12,
          after: opts.after ?? 16,
          line: 245,
        })
      );
    } else {
      blocks.push(
        p([txt(line)], {
          indentLeft: 180,
          before: 12,
          after: opts.after ?? 16,
          line: 245,
        })
      );
    }
  }

  return blocks;
}

function buildStageMetadataTable(
  rows: Array<{ label: string; value: string; boldValue?: boolean }>
): Array<Paragraph | Table> {
  const metaWidth = 6000;
  const metaCols = pctWidths([38, 4, 58], metaWidth);
  return [
    new Table({
      width: { size: metaWidth, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: metaCols,
      rows: rows.map(
        (r) =>
          new TableRow({
            cantSplit: true,
            children: [
              simpleCell(r.label, {
                bold: true,
                size: SZ_META,
                widthTwips: metaCols[0],
                borders: ALL_THIN_BORDERS,
                marginTB: 22,
              }),
              simpleCell(':', {
                bold: true,
                size: SZ_META,
                align: AlignmentType.CENTER,
                widthTwips: metaCols[1],
                borders: ALL_THIN_BORDERS,
                marginTB: 22,
              }),
              simpleCell(r.value, {
                bold: r.boldValue ?? true,
                size: SZ_BODY,
                widthTwips: metaCols[2],
                borders: ALL_THIN_BORDERS,
                marginTB: 22,
              }),
            ],
          })
      ),
    }),
    p([], { before: 20, after: 20, line: 140 }),
  ];
}

function buildSignatoriesAndPlacement(
  placementText: string,
  subtitleText: string
): Array<Paragraph | Table> {
  const sigCols = pctWidths([50, 50]);
  return [
    p([], { before: 50, after: 30 }),
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: [BODY_WIDTH_TWIPS],
      rows: [
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell(placementText, {
              bold: true,
              align: AlignmentType.CENTER,
              size: SZ_BODY,
              widthTwips: BODY_WIDTH_TWIPS,
              borders: ALL_THIN_BORDERS,
              marginTB: 35,
            }),
          ],
        }),
      ],
    }),
    p([txt('Union Bank PLC Unit', { bold: true })], { before: 60, after: 10 }),
    p([txt('Special Asset Management Division', { bold: true })], { before: 0, after: 45 }),
    p([txt(subtitleText, { bold: true })], {
      align: AlignmentType.CENTER,
      before: 35,
      after: 95,
    }),
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: sigCols,
      rows: [
        new TableRow({
          cantSplit: true,
          children: [
            cell(
              [
                p([txt('(Shahidur Rahman khan)', { bold: true })], { before: 180, after: 10 }),
                p([txt('Deputy Managing Director & CRO')], { before: 0, after: 10 }),
                p([txt('Sammilito Islami Bank PLC.')], { before: 0, after: 10 }),
              ],
              { widthTwips: sigCols[0], borders: ALL_THIN_BORDERS, marginLR: 100 }
            ),
            cell(
              [
                p([txt('(Ahmed Zafrul Hasan)', { bold: true })], { before: 180, after: 10 }),
                p([txt('Additional Managing Director & COO')], { before: 0, after: 10 }),
                p([txt('Sammilito Islami Bank PLC.')], { before: 0, after: 10 }),
              ],
              { widthTwips: sigCols[1], borders: ALL_THIN_BORDERS, marginLR: 100 }
            ),
          ],
        }),
      ],
    }),
    p([txt('(Abedur Rahman Sikder)', { bold: true })], {
      align: AlignmentType.CENTER,
      before: 220,
      after: 10,
    }),
    p([txt('Managing Director')], { align: AlignmentType.CENTER, before: 0, after: 10 }),
    p([txt('Sammilito Islami Bank PLC.')], { align: AlignmentType.CENTER, before: 0, after: 10 }),
    p([txt('(Transferee of the assets and liabilities of')], {
      align: AlignmentType.CENTER,
      before: 0,
      after: 10,
    }),
    p([txt('EXIM Bank, SIBL, FSIB, Global Islami')], {
      align: AlignmentType.CENTER,
      before: 0,
      after: 10,
    }),
    p([txt('Bank and Union Bank)')], { align: AlignmentType.CENTER, before: 0, after: 10 }),
  ];
}

function buildSanctionLetterBlocks(stageMemo: ResolvedStageMemo): Array<Paragraph | Table> {
  const { icMemo, stageInputs, stageSections } = stageMemo;
  const ph = icMemo.placeholders;
  const bp = icMemo.branchProposal;
  const v = (code: string) => valOrToken(ph, code);
  const clientName = bp?.section10Particulars?.nameOfClient?.trim() || v('#2#');

  const boardOrder = stageInputs.boardMemoNumber
    ? formatChronologicOrderWithYear(stageInputs.boardMemoNumber, '')
    : '[Chronologic Order of the Board Memo]/26';
  const mecOrder = stageInputs.mecMemoNumber
    ? formatChronologicOrderWithYear(stageInputs.mecMemoNumber, '')
    : '—';
  const icOrder = stageInputs.icMemoNumber
    ? formatChronologicOrderWithYear(stageInputs.icMemoNumber, '')
    : '—';

  const blocks: Array<Paragraph | Table> = [];

  // Reference & Sanction Chain Table
  blocks.push(
    ...buildStageMetadataTable([
      {
        label: 'Sanction Ref. No.',
        value: stageInputs.sanctionRefNumber || '[Sanction Reference Pending]',
        boldValue: Boolean(stageInputs.sanctionRefNumber),
      },
      {
        label: 'Sanction Date',
        value: stageInputs.sanctionDate || stageInputs.boardMemoDate || v('#1#'),
      },
      {
        label: 'Board Memo Reference',
        value: `${boardOrder} dated ${stageInputs.boardMemoDate || '[Board Memo Date]'}`,
      },
      {
        label: 'Prior Committee Chain',
        value: `MEC: ${mecOrder} (${stageInputs.mecMemoDate || '—'}) | IC: ${icOrder} (${stageInputs.icMemoDate || '—'})`,
      },
      {
        label: 'Branch Proposal Reference',
        value: `${v('#64#')} dated ${v('#65#')}`,
      },
      {
        label: 'From',
        value: 'Special Asset Management Division (SAMD), Head Office',
      },
      {
        label: 'To',
        value: `The Head of Branch, ${v('#3#')} Branch`,
      },
    ])
  );

  // Addressee & Formal Subject Banner
  blocks.push(
    p([txt('The Head of Branch,', { bold: true })], { before: 20, after: 5 }),
    p([txt('Sammilito Islami Bank PLC.')], { before: 0, after: 5 }),
    p([txt(`${v('#3#')} Branch`)], { before: 0, after: 25 }),
    p([txt('Dear Sir / Madam,')], { before: 10, after: 20 })
  );

  blocks.push(
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: [BODY_WIDTH_TWIPS],
      rows: [
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell(
              `Subject: Head Office Sanction of Full & Final Exit with Waiver Facility — A/C: ${v('#2#')} (CIS: ${v('#69#')}), Client: ${clientName}, Branch: ${v('#3#')} under BRPD Circular Letter No. 23 dated 29.06.2026.`,
              {
                bold: true,
                size: SZ_BODY,
                widthTwips: BODY_WIDTH_TWIPS,
                borders: ALL_THIN_BORDERS,
                marginTB: 35,
              }
            ),
          ],
        }),
      ],
    }),
    p([txt('[All Figures in this Sanction Letter are in Million BDT Unless Stated Otherwise]', { bold: true })], {
      align: AlignmentType.RIGHT,
      before: 20,
      after: 25,
    })
  );

  // 1. Board Approval Reference & Sanction Basis
  blocks.push(
    ...renderNarrativeParagraphs(
      '1. Board Approval Reference & Sanction Basis:',
      stageSections.sanctionBoardAuthority,
      { before: 35 }
    )
  );

  // 2. Particulars of Sanctioned Facility, Client & Liability Matrix
  const facCols = pctWidths([34, 66]);
  const facilityRows: Array<[string, string, boolean]> = [
    ['Name of the Client & A/C No.', `${clientName} | A/C: ${v('#2#')} (CIS: ${v('#69#')})`, true],
    ['Branch Name & Nature of Business', `${v('#3#')} Branch | ${v('#54#')}`, false],
    ['Mode / Product & Original Limit', `${v('#5#')} | Original Limit: Tk. ${v('#6#')} M (Sec. 2 Limit: Tk. ${v('#56#')} M; Validity: ${v('#58#')})`, false],
    [`NET Outstanding as on ${v('#68#')}`, `Tk. ${v('#25#')} Million (Principal: Tk. ${v('#7#')} M; Charges: Tk. ${v('#8#')} M; Charged Profit: Tk. ${v('#10#')} M; Comp.: Tk. ${v('#11#')} M; Unapplied Comp.: Tk. ${v('#12#')} M)`, true],
    ['Overdue & Cumulative Recovery', `Overdue: Tk. ${v('#59#')} Million | Recovery Till Date: Tk. ${v('#60#')} Million`, false],
    ['Cost of Fund & Permissible Ceiling', `COF Surplus/Deficit: Tk. ${v('#31#')} M | Breakup Total (#7#+#8#+#31#): Tk. ${v('#32#')} M | Max Waiver Considering COF: Tk. ${v('#44#')} M`, false],
    ['Sanctioned Down Payment', `Tk. ${v('#35#')} Million deposited in A/C: ${v('#36#')} on ${v('#68#')}`, true],
    ['Sanctioned Waiver Amount', `Tk. ${v('#4#')} Million (plus accrued/unapplied profit up to execution date)`, true],
    ['Sanctioned Settlement / Exit Amount', `Tk. ${v('#27#')} Million (Lump Sum without charging further profit)`, true],
    ['Sanctioned Settlement Period', `${v('#28#')}`, true],
    ['Verified Source of Repayment', `${icMemo.sourceOfRepayment || 'As stipulated in the approved Board Memorandum'}`, false],
  ];

  blocks.push(
    p(
      [
        txt('2. Particulars of Sanctioned Facility, Client & Liability Matrix:', {
          bold: true,
          underline: true,
        }),
      ],
      { before: 45, after: 15 }
    ),
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: facCols,
      rows: [
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            simpleCell('Sanction Parameter', {
              bold: true,
              widthTwips: facCols[0],
              borders: ALL_THIN_BORDERS,
              marginTB: 22,
            }),
            simpleCell('Sanctioned / Cross-Verified Particular', {
              bold: true,
              widthTwips: facCols[1],
              borders: ALL_THIN_BORDERS,
              marginTB: 22,
            }),
          ],
        }),
        ...facilityRows.map(
          ([lbl, val, isBold]) =>
            new TableRow({
              cantSplit: true,
              children: [
                simpleCell(lbl, {
                  bold: true,
                  widthTwips: facCols[0],
                  borders: ALL_THIN_BORDERS,
                  marginTB: 22,
                }),
                simpleCell(val, {
                  bold: isBold,
                  widthTwips: facCols[1],
                  borders: ALL_THIN_BORDERS,
                  marginTB: 22,
                }),
              ],
            })
        ),
      ],
    })
  );

  // 3. AI Cross-Verified Facility & Liability Confirmation
  blocks.push(
    ...renderNarrativeParagraphs(
      '3. AI Cross-Verified Facility & Liability Confirmation:',
      stageSections.sanctionCrossVerifiedSummary,
      { before: 45 }
    )
  );

  // 4. Collateral Security, Legal & Regulatory Clearance Schedule
  blocks.push(
    ...renderNarrativeParagraphs(
      '4. Collateral Security, Legal & Regulatory Clearance Schedule:',
      stageSections.sanctionSecurityAndCompliance,
      { before: 45 }
    )
  );

  // 5. Operative Terms, Conditions & Redemption Covenants
  blocks.push(
    ...renderNarrativeParagraphs(
      '5. Operative Terms, Conditions & Redemption Covenants:',
      stageSections.sanctionSpecialConditions,
      { before: 45 }
    )
  );

  const mandatoryClauses: Array<[string, string, boolean]> = [
    [
      'vii.',
      'The waiver of profit/compensation shall be adjusted from FAD, Head Office from the provision amount kept (if any), and the remaining amount shall be adjusted gradually by the Branch upon full realization of the sanctioned settlement amount within the approved deadline.',
      false,
    ],
    [
      'viii.',
      'Branch shall strictly comply with Bangladesh Bank BRPD Circular No. 06/22, Circular Letter No. 18/22, 46/22, BRPD Circular Letter No. 23 dated 29.06.2026 and all applicable Head Office instructions.',
      false,
    ],
    [
      'ix.',
      'The Profit rate of the stipulated investment shall NOT be changed or reduced to 0.00% until full and final adjustment of the Investment, as approved by Head Office, and all legal/suit withdrawal charges shall be borne by the client.',
      true,
    ],
  ];
  for (const [num, clause, isBold] of mandatoryClauses) {
    blocks.push(
      p([txt(`${num}  `, { bold: isBold }), txt(clause, { bold: isBold })], {
        indentLeft: 320,
        before: 18,
        after: 18,
        line: 245,
      })
    );
  }

  blocks.push(
    ...buildSignatoriesAndPlacement(
      'Issued under Head Office Sanction Authority. Strict compliance with all sanctioned terms, timelines, and BRPD directives is mandatory.',
      'Conveyed for immediate execution, recovery realization, and compliance by the Branch.'
    )
  );

  return blocks;
}

export async function generateStageMemoDocxBlob(stageMemo: ResolvedStageMemo): Promise<Blob> {
  const { docType, icMemo, stageInputs, stageSections } = stageMemo;
  const ph = icMemo.placeholders;
  const v = (code: string) => valOrToken(ph, code);

  const badgeLabel =
    docType === 'mec'
      ? 'MEC Memo'
      : docType === 'board'
        ? 'Board Memo'
        : 'Sanction Letter';

  const { header, footer } = await buildBankHeaderAndFooter(badgeLabel);
  const body: Array<Paragraph | Table> = [];

  if (docType === 'mec') {
    const icOrderDisplay = stageInputs.icMemoNumber
      ? formatChronologicOrderWithYear(stageInputs.icMemoNumber, '')
      : '';
    const mecOrderDisplay = stageInputs.mecMemoNumber
      ? formatChronologicOrderWithYear(stageInputs.mecMemoNumber, '')
      : '';

    body.push(
      ...buildStageMetadataTable([
        { label: 'Dated', value: stageInputs.mecMemoDate || stageInputs.icMemoDate || v('#1#') },
        {
          label: 'Chronologic Order of the MEC Memo',
          value: mecOrderDisplay,
          boldValue: Boolean(mecOrderDisplay),
        },
        {
          label: 'Chronologic Order of the IC Memo',
          value: icOrderDisplay,
          boldValue: Boolean(icOrderDisplay),
        },
        {
          label: 'IC Memo Date',
          value: stageInputs.icMemoDate || '',
          boldValue: Boolean(stageInputs.icMemoDate),
        },
        { label: 'To', value: 'Management Executive Committee (MEC)' },
        { label: 'From', value: 'SAMD' },
      ])
    );

    const sec16Text =
      stageSections.mecIcRecommendation.trim() ||
      buildMecIcRecommendationText(stageInputs.icMemoNumber, stageInputs.icMemoDate);

    const extraBlocksAfterSamd = renderCommitteeRecommendationBlocks(
      '16. Recommendation of the Investment Committee (IC):',
      sec16Text,
      { before: 45, after: 18 }
    );

    // Full retained IC Memo core blocks (Sections 1-15 -> Sec 16 IC Rec -> 17. Particulars -> 18. Special Terms)
    body.push(
      ...buildICMemoCoreBlocks(icMemo, {
        includeClosingBlock: false,
        skipMetadataTable: true,
        facilityStartNumber: 17,
        extraBlocksAfterSamd,
      })
    );

    body.push(
      ...renderNarrativeParagraphs(
        '19. Executive Appraisal & Waiver Rationale for MEC:',
        stageSections.mecExecutiveSummary,
        { before: 45 }
      ),
      ...renderNarrativeParagraphs(
        '20. Observation and Recommendation of SAMD for MEC:',
        stageSections.mecSamdRecommendation,
        { before: 40 }
      ),
      ...buildSignatoriesAndPlacement(
        'The proposal is hereby placed before the Management Executive Committee (MEC) for kind recommendation/approval.',
        'Submitted before the Management Executive Committee (MEC) for kind decision.'
      )
    );
  } else if (docType === 'board') {
    const icOrderDisplay = stageInputs.icMemoNumber
      ? formatChronologicOrderWithYear(stageInputs.icMemoNumber, '')
      : '—';
    const mecOrderDisplay = stageInputs.mecMemoNumber
      ? formatChronologicOrderWithYear(stageInputs.mecMemoNumber, '')
      : '';
    const boardOrderDisplay = stageInputs.boardMemoNumber
      ? formatChronologicOrderWithYear(stageInputs.boardMemoNumber, '')
      : '';

    body.push(
      ...buildStageMetadataTable([
        { label: 'Dated', value: stageInputs.boardMemoDate || stageInputs.mecMemoDate || v('#1#') },
        {
          label: 'Chronologic Order of the Board Memo',
          value: boardOrderDisplay,
          boldValue: Boolean(boardOrderDisplay),
        },
        {
          label: 'Chronologic Order of the MEC Memo',
          value: mecOrderDisplay,
          boldValue: Boolean(mecOrderDisplay),
        },
        {
          label: 'MEC Memo Date',
          value: stageInputs.mecMemoDate || '',
          boldValue: Boolean(stageInputs.mecMemoDate),
        },
        {
          label: 'Chronologic Order of the IC Memo & Date',
          value: `${icOrderDisplay} dated ${stageInputs.icMemoDate || '—'}`,
        },
        { label: 'To', value: 'Board of Directors' },
        { label: 'From', value: 'Management Executive Committee (MEC) / SAMD' },
      ])
    );

    const carriedSec16Text =
      stageSections.mecIcRecommendation.trim() ||
      (stageInputs.icMemoNumber || stageInputs.icMemoDate
        ? buildMecIcRecommendationText(stageInputs.icMemoNumber, stageInputs.icMemoDate)
        : '');

    const sec17Text =
      stageSections.boardMecRecommendation.trim() ||
      buildBoardMecRecommendationText(stageInputs.mecMemoNumber, stageInputs.mecMemoDate);

    const extraBlocksAfterSamd: Array<Paragraph | Table> = [];
    if (carriedSec16Text) {
      extraBlocksAfterSamd.push(
        ...renderCommitteeRecommendationBlocks(
          '16. Recommendation of the Investment Committee (IC):',
          carriedSec16Text,
          { before: 45, after: 16 }
        )
      );
    }
    extraBlocksAfterSamd.push(
      ...renderCommitteeRecommendationBlocks(
        '17. Management Executive Committee (MEC) Recommendation:',
        sec17Text,
        { before: 40, after: 18 }
      )
    );

    body.push(
      ...buildICMemoCoreBlocks(icMemo, {
        includeClosingBlock: false,
        skipMetadataTable: true,
        facilityStartNumber: 18,
        extraBlocksAfterSamd,
      })
    );

    if (stageSections.mecExecutiveSummary.trim()) {
      body.push(
        ...renderNarrativeParagraphs(
          'Executive Appraisal & Waiver Rationale (Carried Forward from MEC Memo):',
          stageSections.mecExecutiveSummary,
          { before: 45 }
        )
      );
    }
    if (stageSections.mecSamdRecommendation.trim()) {
      body.push(
        ...renderNarrativeParagraphs(
          'Observation and Recommendation of SAMD for MEC (Carried Forward):',
          stageSections.mecSamdRecommendation,
          { before: 35 }
        )
      );
    }

    body.push(
      ...renderNarrativeParagraphs(
        '20. Board Memorandum Summary & Financial Justification:',
        stageSections.boardProposalRationale,
        { before: 45 }
      ),
      ...renderNarrativeParagraphs(
        '21. Proposed Board Resolution & Approval Terms:',
        stageSections.boardResolutionTerms,
        { before: 40 }
      ),
      ...buildSignatoriesAndPlacement(
        'Submitted before the Honourable Board of Directors for kind consideration and approval.',
        'Submitted for kind approval of the Board of Directors.'
      )
    );
  } else {
    body.push(...buildSanctionLetterBlocks(stageMemo));
  }

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            size: {
              width: A4_WIDTH_TWIPS,
              height: A4_HEIGHT_TWIPS,
            },
            margin: {
              top: 1520,
              bottom: 680,
              left: MARGIN_LR_TWIPS,
              right: MARGIN_LR_TWIPS,
              header: 300,
              footer: 280,
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
