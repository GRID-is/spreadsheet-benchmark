import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { fixturePath } from './fixture.ts';

test('fixturePath keeps xlsx names and maps xls to xlsx', () => {
  // Arrange
  const expectations = [
    [ 'mathematical/abs.xlsx', 'mathematical/abs.xlsx' ],
    [ 'eval/FormulaEvalTestData.xls', 'eval/FormulaEvalTestData.xlsx' ],
    [ 'functions/57003-FixedFunctionTestCaseData.xls', 'functions/57003-FixedFunctionTestCaseData.xlsx' ],
    [ 'xls/inside.xlsx', 'xls/inside.xlsx' ],
  ];

  // Act & Assert
  for (const [ file, fixture ] of expectations) {
    assert.equal(fixturePath(file), fixture, file);
  }
});

test('no two workbooks of a suite share a fixture', () => {
  const suites = new URL('../suites/', import.meta.url);
  for (const suite of readdirSync(suites)) {
    const manifest = new URL(`${suite}/manifest.json`, suites);
    if (!existsSync(manifest)) {
      continue;
    }
    const files: Array<{ file: string }> = JSON.parse(readFileSync(manifest, 'utf8')).files;
    const byFixture = new Map<string, string[]>();
    for (const { file } of files) {
      byFixture.set(fixturePath(file), [ ...byFixture.get(fixturePath(file)) ?? [], file ]);
    }
    const shared = [ ...byFixture.values() ].filter(group => group.length > 1);
    assert.deepEqual(shared, [], `the workbooks of suite ${suite} have a fixture each`);
  }
});
