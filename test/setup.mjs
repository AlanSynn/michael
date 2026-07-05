// Test bootstrap for pure-module unit tests (node --test).
// Provides the webpack DefinePlugin globals that src/shared/zai.js reads at
// module-eval time, so the Z.AI helpers can be imported without webpack.
globalThis.__ZAI_API_KEY__ = "test-key-do-not-use";
globalThis.__ZAI_CODING_BASE_URL__ = "https://api.z.ai/api/coding/paas/v4";
