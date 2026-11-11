import type { Adapter } from '../../src/adapter.ts';

const adapter: Adapter = {
  isCase: (sheetIndex, ref) => sheetIndex > 0 && /^A\d+$/.test(ref),
  upstreamRef: ref => ref.replace(/^A/, 'B'),
};

export default adapter;
