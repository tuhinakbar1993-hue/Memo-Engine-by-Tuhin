import type { MemoEngineDefinition } from '../types/memo';

export const MEMO_ENGINES: MemoEngineDefinition[] = [
  {
    id: 'ic',
    name: 'IC Memo',
    status: 'implemented',
    description: 'Investment Committee Memorandum for Exit with Waiver under BRPD Circular Letter No. 23.',
  },
  {
    id: 'mec',
    name: 'MEC Memo',
    status: 'coming_soon',
    description: 'Management Executive Committee Memorandum rules will be integrated here.',
  },
  {
    id: 'board',
    name: 'Board Memo',
    status: 'coming_soon',
    description: 'Board of Directors Memorandum rules will be integrated here.',
  },
  {
    id: 'sanction',
    name: 'Sanction',
    status: 'coming_soon',
    description: 'Head Office Sanction Letter rules will be integrated here.',
  },
];
