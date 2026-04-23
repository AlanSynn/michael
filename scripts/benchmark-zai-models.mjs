#!/usr/bin/env node

const BASE_URL = (process.env.ZAI_CODING_BASE_URL || "https://api.z.ai/api/coding/paas/v4").replace(/\/+$/, "");
const API_KEY = process.env.ZAI_API_KEY || "";
const DEFAULT_MODELS = [
  "glm-5-turbo",
  "glm-4.7-flash",
  "glm-4.7-flashx",
  "glm-4.5-flash",
  "glm-4.5-air",
  "glm-4.5-airx",
  "glm-4.7",
  "glm-5.1",
];
const MODELS = (process.env.ZAI_BENCH_MODELS || DEFAULT_MODELS.join(","))
  .split(",")
  .map((model) => model.trim())
  .filter(Boolean);
const ROUNDS = Number.parseInt(process.env.ZAI_BENCH_ROUNDS || "2", 10);
const TIMEOUT_MS = Number.parseInt(process.env.ZAI_BENCH_TIMEOUT_MS || "180000", 10);

const SCENARIOS = [
  {
    name: "short",
    maxTokens: 256,
    prompt: "In Korean, summarize why email assistants should keep user-configurable prompts. Answer in 5 bullets.",
  },
  {
    name: "long",
    maxTokens: 2048,
    prompt:
      "Write a Korean technical memo for an Outlook add-in migration from Gemini to Z.AI GLM Coding Plan. Include architecture, authentication, model dropdown refresh, prompt templates, risks, test plan, and rollback. Use detailed sections and enough substance to measure long-output throughput.",
  },
];

function usage() {
  console.error("Missing ZAI_API_KEY. Run: ZAI_API_KEY=... node scripts/benchmark-zai-models.mjs");
}

function nowMs() {
  return Number(process.hrtime.bigint()) / 1_000_000;
}

function estimateTokens(text) {
  return Math.max(1, Math.round(text.length / 4));
}

function parseSsePayload(line) {
  if (!line.startsWith("data:")) return null;
  const payload = line.slice(5).trim();
  if (!payload || payload === "[DONE]") return null;
  try {
    return JSON.parse(payload);
  } catch {
    return null;
  }
}

async function chatStream(model, scenario) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), TIMEOUT_MS);
  const startMs = nowMs();
  let firstTokenMs = null;
  let output = "";

  try {
    const response = await fetch(`${BASE_URL}/chat/completions`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${API_KEY}`,
        "Content-Type": "application/json",
      },
      signal: controller.signal,
      body: JSON.stringify({
        model,
        stream: true,
        thinking: { type: "disabled" },
        temperature: 0.2,
        max_tokens: scenario.maxTokens,
        messages: [{ role: "user", content: scenario.prompt }],
      }),
    });

    if (!response.ok) {
      const body = await response.text().catch(() => "");
      throw new Error(`HTTP ${response.status}: ${body.slice(0, 300)}`);
    }

    const reader = response.body?.getReader();
    if (!reader) throw new Error("No response body stream.");

    const decoder = new TextDecoder();
    let buffer = "";
    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split(/\r?\n/);
      buffer = lines.pop() || "";
      for (const line of lines) {
        const data = parseSsePayload(line.trim());
        const delta = data?.choices?.[0]?.delta;
        const chunk = delta?.content || delta?.reasoning_content || "";
        if (chunk) {
          if (firstTokenMs === null) firstTokenMs = nowMs() - startMs;
          output += chunk;
        }
      }
    }

    const totalMs = nowMs() - startMs;
    const tokens = estimateTokens(output);
    return {
      model,
      scenario: scenario.name,
      ok: true,
      ttftMs: Math.round(firstTokenMs ?? totalMs),
      totalMs: Math.round(totalMs),
      outputChars: output.length,
      estimatedTokens: tokens,
      estimatedTokensPerSec: Number((tokens / (totalMs / 1000)).toFixed(2)),
    };
  } catch (error) {
    const totalMs = nowMs() - startMs;
    return {
      model,
      scenario: scenario.name,
      ok: false,
      totalMs: Math.round(totalMs),
      error: error?.name === "AbortError" ? `timeout after ${TIMEOUT_MS}ms` : error.message,
    };
  } finally {
    clearTimeout(timeout);
  }
}

function median(values) {
  const sorted = values.slice().sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
}

function summarize(results) {
  const rows = [];
  for (const model of MODELS) {
    for (const scenario of SCENARIOS) {
      const group = results.filter((result) => result.model === model && result.scenario === scenario.name);
      const ok = group.filter((result) => result.ok);
      const failed = group.filter((result) => !result.ok);
      rows.push({
        model,
        scenario: scenario.name,
        success: `${ok.length}/${group.length}`,
        medianTtftMs: ok.length ? Math.round(median(ok.map((result) => result.ttftMs))) : null,
        medianTotalMs: ok.length ? Math.round(median(ok.map((result) => result.totalMs))) : null,
        medianTokensPerSec: ok.length
          ? Number(median(ok.map((result) => result.estimatedTokensPerSec)).toFixed(2))
          : null,
        errors: failed.map((result) => result.error).join(" | "),
      });
    }
  }
  return rows;
}

async function main() {
  if (!API_KEY) {
    usage();
    process.exit(2);
  }

  console.log(JSON.stringify({ baseUrl: BASE_URL, models: MODELS, rounds: ROUNDS, scenarios: SCENARIOS.map((s) => s.name) }, null, 2));
  const results = [];
  for (const model of MODELS) {
    for (const scenario of SCENARIOS) {
      for (let round = 1; round <= ROUNDS; round += 1) {
        process.stderr.write(`benchmark ${model} ${scenario.name} round ${round}/${ROUNDS}\n`);
        const result = await chatStream(model, scenario);
        results.push(result);
        console.log(JSON.stringify(result));
      }
    }
  }

  const summary = summarize(results);
  console.log("\nSUMMARY");
  console.table(summary);

  const longWinners = summary
    .filter((row) => row.scenario === "long" && row.medianTotalMs !== null)
    .sort((a, b) => a.medianTotalMs - b.medianTotalMs);
  const shortWinners = summary
    .filter((row) => row.scenario === "short" && row.medianTtftMs !== null)
    .sort((a, b) => a.medianTtftMs - b.medianTtftMs);

  console.log("\nWINNERS");
  console.log(JSON.stringify({ fastestShortTtft: shortWinners[0] || null, fastestLongTotal: longWinners[0] || null }, null, 2));
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
