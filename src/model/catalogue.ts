export type ApiaryStance = 'supported' | 'unsupported' | 'never';

export type FunctionEntry = {
  name: string,
  inExcel: boolean,
  apiary: ApiaryStance,
  legacy: boolean,
  nondeterministic: boolean,
};

export type Catalogue = {
  source: { apiaryVersion: string, generatedAt: string },
  functions: Record<string, FunctionEntry>,
};

export function lookup (catalogue: Catalogue, name: string): FunctionEntry {
  return catalogue.functions[name] ?? { name, inExcel: false, apiary: 'unsupported', legacy: false, nondeterministic: false };
}
