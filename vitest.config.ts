import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    globals: true,
    coverage: {
      exclude: [
        "src/cli.ts", // CLI command wiring — tested via manual smoke tests
        "src/client.ts", // Preconfigured fetch clients — thin config, no logic
      ],
    },
  },
});
