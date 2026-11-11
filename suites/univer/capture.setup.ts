import { appendFileSync } from 'node:fs';
import { expect } from 'vitest';
import type { BaseFunction } from './src/functions/base-function';
import { ALL_IMPLEMENTED_FUNCTIONS } from './src/functions';
import type { BaseValueObject } from './src/engine/value-object/base-value-object';
import type { ArrayValueObject } from './src/engine/value-object/array-value-object';

type Scalar = string | number | boolean | null;
type Arg = { kind: 'scalar', value: Scalar } | { kind: 'empty' } | { kind: 'error', value: string } | { kind: 'array', rows: Scalar[][] } | { kind: 'unsupported', what: string };

const out = `${process.env['UNIVER_CAPTURE_OUT']}.${process.pid}.jsonl`;
let depth = 0;

function scalar (v: BaseValueObject): Scalar {
  const value = v.getValue();
  return typeof value === 'number' || typeof value === 'string' || typeof value === 'boolean' ? value : null;
}

function serialize (v: unknown): Arg {
  if (v === null || v === undefined) {
    return { kind: 'empty' };
  }
  const obj = v as BaseValueObject & { isReferenceObject?: () => boolean, toArrayValueObject?: () => ArrayValueObject };
  if (typeof obj.isReferenceObject === 'function' && obj.isReferenceObject() && typeof obj.toArrayValueObject === 'function') {
    return serialize(obj.toArrayValueObject());
  }
  if (typeof obj.isNull !== 'function') {
    return { kind: 'unsupported', what: obj?.constructor?.name ?? typeof obj };
  }
  if (obj.isNull()) {
    return { kind: 'empty' };
  }
  if (obj.isError()) {
    return { kind: 'error', value: String(obj.getValue()) };
  }
  if (obj.isLambda() || obj.isCube?.()) {
    return { kind: 'unsupported', what: obj.constructor.name };
  }
  if (obj.isArray()) {
    const rows = (obj as ArrayValueObject).getArrayValue().map(row => (row ?? []).map(cell => cell == null || cell.isNull() ? null : cell.isError() ? String(cell.getValue()) : scalar(cell)));
    return { kind: 'array', rows };
  }
  return { kind: 'scalar', value: scalar(obj) };
}

for (const [ Ctor ] of ALL_IMPLEMENTED_FUNCTIONS) {
  const original = Ctor.prototype.calculate;
  Ctor.prototype.calculate = function (this: BaseFunction, ...args: BaseValueObject[]) {
    if (depth > 0) {
      return original.apply(this, args);
    }
    depth++;
    try {
      const result = original.apply(this, args);
      const state = expect.getState();
      appendFileSync(out, JSON.stringify({
        file: state.testPath,
        test: state.currentTestName,
        name: this.name,
        args: args.map(serialize),
        result: serialize(result),
      }) + '\n');
      return result;
    }
    finally {
      depth--;
    }
  };
}
