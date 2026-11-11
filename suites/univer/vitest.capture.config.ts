import { mergeConfig } from 'vitest/config';
import createConfig from '@univerjs-infra/shared/vitest';

export default mergeConfig(createConfig(), {
  test: {
    setupFiles: [ './capture.setup.ts' ],
    include: [ 'src/functions/**/__tests__/**/*.spec.ts' ],
    coverage: { enabled: false },
  },
});
