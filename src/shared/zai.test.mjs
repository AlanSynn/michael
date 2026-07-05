import { test } from "node:test";
import assert from "node:assert/strict";

// zai.js is plain ESM, so Node's native runner imports it directly.
const {
  dedupeModels,
  extractModelNames,
  normalizeModelName,
  normalizeTextContent,
  fetchAvailableModels,
  generateText,
  getDefaultZaiModels,
  getDefaultZaiModel,
} = await import("./zai.js");

// --- pure helpers -----------------------------------------------------------

test("normalizeModelName trims + lowercases, non-strings become empty", () => {
  assert.equal(normalizeModelName("  GLM-4.5  "), "glm-4.5");
  assert.equal(normalizeModelName(undefined), "");
  assert.equal(normalizeModelName(42), "");
});

test("dedupeModels is case-insensitive and drops empties", () => {
  assert.deepEqual(dedupeModels(["glm-4.5", "GLM-4.5", "  ", "glm-4.6"]), ["glm-4.5", "glm-4.6"]);
  assert.deepEqual(dedupeModels([]), []);
});

test("normalizeTextContent handles string, array-of-parts, and garbage", () => {
  assert.equal(normalizeTextContent("  hi "), "hi");
  assert.equal(
    normalizeTextContent([{ text: "a" }, { text: "b" }]),
    "ab"
  );
  assert.equal(normalizeTextContent(["x", "y"]), "xy");
  assert.equal(normalizeTextContent(null), "");
});

test("extractModelNames pulls glm-* ids out of nested /models payloads", () => {
  const payload = {
    object: "list",
    data: [
      { id: "glm-4.5-air", object: "model" },
      { id: "glm-4.6", object: "model" },
      { id: "embedding-2", object: "model" }, // not a glm-* model -> ignored
      { id: "GLM-4.5-AIR" }, // case dup of first -> collapsed
    ],
  };
  assert.deepEqual(extractModelNames(payload), ["glm-4.5-air", "glm-4.6"]);
});

test("extractModelNames returns [] when no glm-* models are present", () => {
  assert.deepEqual(extractModelNames({ data: [{ id: "embedding-2" }] }), []);
  assert.deepEqual(extractModelNames(null), []);
});

// --- fetch-backed paths (globalThis.fetch mocked) ---------------------------

function mockFetch(responder) {
  const calls = [];
  globalThis.fetch = async (url, options) => {
    calls.push({ url: String(url), options });
    return responder({ url: String(url), options });
  };
  return calls;
}

test("fetchAvailableModels discovers + dedupes glm-* models from /models", async () => {
  const calls = mockFetch(() => ({
    ok: true,
    json: async () => ({
      data: [
        { id: "glm-4.5-air" },
        { id: "glm-4.5-air" },
        { id: "glm-4.6" },
        { id: "text-embedding-2" },
      ],
    }),
  }));

  const models = await fetchAvailableModels({ apiKey: "key" });
  assert.deepEqual(models, ["glm-4.5-air", "glm-4.6"]);
  assert.match(calls[0].url, /\/models$/);
  assert.equal(calls[0].options.headers.Authorization, "Bearer key");
});

test("fetchAvailableModels surfaces the provider error message on failure", async () => {
  mockFetch(() => ({
    ok: false,
    status: 401,
    json: async () => ({ error: { message: "Invalid API key" } }),
  }));

  await assert.rejects(() => fetchAvailableModels({ apiKey: "bad" }), /Invalid API key/);
});

test("fetchAvailableModels throws when the provider returns no glm-* models", async () => {
  mockFetch(() => ({
    ok: true,
    json: async () => ({ data: [{ id: "embedding-2" }] }),
  }));

  await assert.rejects(
    () => fetchAvailableModels({ apiKey: "key" }),
    /No models were returned by the provider/
  );
});

test("generateText posts to /chat/completions and returns trimmed content", async () => {
  const calls = mockFetch(() => ({
    ok: true,
    json: async () => ({
      choices: [{ message: { content: "  translated body  " } }],
    }),
  }));

  const text = await generateText("translate this", {
    apiKey: "key",
    model: "glm-4.5-air",
    temperature: 0.3,
  });

  assert.equal(text, "translated body");
  assert.match(calls[0].url, /\/chat\/completions$/);
  const body = JSON.parse(calls[0].options.body);
  assert.equal(body.model, "glm-4.5-air");
  assert.equal(body.messages[0].content, "translate this");
  assert.equal(body.temperature, 0.3);
});

test("generateText handles array message content parts", async () => {
  mockFetch(() => ({
    ok: true,
    json: async () => ({
      choices: [{ message: { content: [{ text: "part-1-" }, { text: "part-2" }] } }],
    }),
  }));

  const text = await generateText("p", { apiKey: "key" });
  assert.equal(text, "part-1-part-2");
});

test("generateText rejects with the provider error on non-ok response", async () => {
  mockFetch(() => ({
    ok: false,
    status: 429,
    json: async () => ({ error: { message: "Rate limited" } }),
  }));

  await assert.rejects(
    () => generateText("p", { apiKey: "key" }),
    /Rate limited/
  );
});

test("defaults are populated and defensive copies", () => {
  assert.ok(getDefaultZaiModels().length >= 1);
  assert.equal(getDefaultZaiModel(), "glm-4.5-air");
  // mutating the returned default list must not affect later calls
  const a = getDefaultZaiModels();
  a.push("glm-fake");
  assert.ok(!getDefaultZaiModels().includes("glm-fake"));
});
