import assert from 'node:assert/strict';
import { test } from 'node:test';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

test('spreadsheetgear says that dotnet does not run when it is missing', t => {
  // A fresh working directory has no built shim, so the fixture's evaluation starts a build.
  const dir = mkdtempSync(join(tmpdir(), 'spreadsheetgear-'));
  t.after(() => rmSync(dir, { recursive: true, force: true }));
  const script = `
    import { engine } from ${JSON.stringify(new URL('./spreadsheetgear.ts', import.meta.url).href)};
    try {
      await engine.buildFile('workbook.xlsx', []);
      console.log('built');
    }
    catch (err) {
      console.log(err.message);
    }
  `;

  const out = execFileSync(process.execPath, [ '--input-type=module', '-e', script ], {
    cwd: dir, env: { ...process.env, DOTNET: join(dir, 'dotnet') }, encoding: 'utf8',
  });

  assert.match(out.trim(), /dotnet does not run \(brew install dotnet\)/, 'the error names the missing dotnet');
});

test('spreadsheetgear reads its version from the project file, without building the shim', t => {
  const dir = mkdtempSync(join(tmpdir(), 'spreadsheetgear-'));
  t.after(() => rmSync(dir, { recursive: true, force: true }));
  const script = `
    import { engine } from ${JSON.stringify(new URL('./spreadsheetgear.ts', import.meta.url).href)};
    console.log(engine.version);
  `;

  const out = execFileSync(process.execPath, [ '--input-type=module', '-e', script ], {
    cwd: dir, env: { ...process.env, DOTNET: join(dir, 'dotnet') }, encoding: 'utf8',
  });

  assert.match(out.trim(), /^\d+\.\d+\.\d+$/, 'the version is read with no dotnet to build with');
});
