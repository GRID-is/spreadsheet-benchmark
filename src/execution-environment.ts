import { existsSync } from 'node:fs';
import { hostname } from 'node:os';

export type ExecutionEnvironment = {
  machine?: string,
  platform: string,
  architecture: string,
  node: string,
  locale: string,
  timeZone: string,
  variables: Record<string, string>,
  image?: string,
  sourceRevision?: string,
};

export function executionEnvironment (): ExecutionEnvironment {
  const container = Boolean(process.env.ENGINE_COMPATIBILITY_ENGINE) || existsSync('/.dockerenv') || existsSync('/run/.containerenv');
  const machine = process.env.ENGINE_COMPATIBILITY_MACHINE || (container ? undefined : hostname());
  const variables: Record<string, string> = {};
  for (const key of [ 'LANG', 'LC_ALL', 'TZ' ]) {
    const value = process.env[key];
    if (value !== undefined) variables[key] = value;
  }
  return {
    ...(machine ? { machine } : {}),
    platform: process.platform,
    architecture: process.arch,
    node: process.version,
    locale: new Intl.NumberFormat().resolvedOptions().locale,
    timeZone: new Intl.DateTimeFormat().resolvedOptions().timeZone,
    variables,
    ...(process.env.ENGINE_COMPATIBILITY_IMAGE === undefined ? {} : { image: process.env.ENGINE_COMPATIBILITY_IMAGE }),
    ...(process.env.ENGINE_COMPATIBILITY_REVISION === undefined ? {} : { sourceRevision: process.env.ENGINE_COMPATIBILITY_REVISION }),
  };
}
