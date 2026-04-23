/* global __ZAI_API_KEY__, __ZAI_CODING_BASE_URL__, AbortController, clearTimeout, fetch, setTimeout */

const DEFAULT_ZAI_BASE_URL = "https://api.z.ai/api/coding/paas/v4";
const DEFAULT_ZAI_MODEL = "glm-5-turbo";
const DEFAULT_ZAI_REPLY_MODEL = "glm-5-turbo";
const MODEL_DISCOVERY_TIMEOUT_MS = 5000;
const CHAT_COMPLETION_TIMEOUT_MS = 120000;
const FALLBACK_ZAI_MODELS = Object.freeze([
  "glm-5-turbo",
  "glm-5.1",
  "glm-5",
  "glm-4.7",
  "glm-4.7-flash",
  "glm-4.7-flashx",
  "glm-4.6",
  "glm-4.5",
  "glm-4.5-air",
  "glm-4.5-x",
  "glm-4.5-airx",
  "glm-4.5-flash",
  "glm-4-32b-0414-128k",
]);

function normalizeModelName(value) {
  return typeof value === "string" ? value.trim().toLowerCase() : "";
}

function dedupeModels(models) {
  const seen = new Set();
  return models.filter((model) => {
    const normalized = normalizeModelName(model);
    if (!normalized || seen.has(normalized)) {
      return false;
    }

    seen.add(normalized);
    return true;
  });
}

function normalizeTextContent(content) {
  if (typeof content === "string") {
    return content.trim();
  }

  if (Array.isArray(content)) {
    return content
      .map((item) => {
        if (typeof item === "string") {
          return item;
        }

        if (item && typeof item.text === "string") {
          return item.text;
        }

        return "";
      })
      .join("")
      .trim();
  }

  return "";
}

function getZaiBaseUrl() {
  const configuredBaseUrl =
    typeof __ZAI_CODING_BASE_URL__ === "string" && __ZAI_CODING_BASE_URL__.trim()
      ? __ZAI_CODING_BASE_URL__.trim()
      : DEFAULT_ZAI_BASE_URL;

  return configuredBaseUrl.replace(/\/+$/, "");
}

function getZaiApiKey() {
  return typeof __ZAI_API_KEY__ === "string" ? __ZAI_API_KEY__.trim() : "";
}

function hasZaiApiKey() {
  return Boolean(getZaiApiKey());
}

function getDefaultZaiModels() {
  return [...FALLBACK_ZAI_MODELS];
}

function getDefaultZaiModel() {
  return DEFAULT_ZAI_MODEL;
}

function getDefaultZaiReplyModel() {
  return DEFAULT_ZAI_REPLY_MODEL;
}

function createTimeoutSignal(timeoutMs) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeoutMs);

  return {
    signal: controller.signal,
    clear() {
      clearTimeout(timeoutId);
    },
  };
}

function extractModelNames(payload) {
  const discovered = [];

  function visit(value) {
    if (!value) {
      return;
    }

    if (Array.isArray(value)) {
      value.forEach(visit);
      return;
    }

    if (typeof value === "string") {
      const normalized = normalizeModelName(value);
      if (normalized.startsWith("glm-")) {
        discovered.push(normalized);
      }
      return;
    }

    if (typeof value !== "object") {
      return;
    }

    ["id", "model", "name"].forEach((field) => {
      if (typeof value[field] === "string") {
        const normalized = normalizeModelName(value[field]);
        if (normalized.startsWith("glm-")) {
          discovered.push(normalized);
        }
      }
    });

    Object.values(value).forEach(visit);
  }

  visit(payload);

  return dedupeModels(discovered);
}

async function fetchJson(url, options = {}, timeoutMs = MODEL_DISCOVERY_TIMEOUT_MS) {
  const timeout = createTimeoutSignal(timeoutMs);

  try {
    const response = await fetch(url, { ...options, signal: timeout.signal });
    const payload = await response.json().catch(() => ({}));

    if (!response.ok) {
      throw new Error(payload?.error?.message || `Request failed (${response.status})`);
    }

    return payload;
  } catch (error) {
    if (error?.name === "AbortError") {
      throw new Error(`Request timed out after ${Math.ceil(timeoutMs / 1000)} seconds.`);
    }

    throw error;
  } finally {
    timeout.clear();
  }
}

async function generateText(prompt, options = {}) {
  const apiKey = options.apiKey || getZaiApiKey();
  const model = normalizeModelName(options.model) || DEFAULT_ZAI_MODEL;

  if (!apiKey) {
    throw new Error("ZAI_API_KEY is not configured.");
  }

  const payload = await fetchJson(
    `${getZaiBaseUrl()}/chat/completions`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model,
        messages: [
          {
            role: "user",
            content: prompt,
          },
        ],
        thinking: {
          type: "disabled",
        },
        temperature: options.temperature ?? 0.4,
        max_tokens: options.maxTokens ?? 4096,
        stream: false,
      }),
    },
    options.timeoutMs ?? CHAT_COMPLETION_TIMEOUT_MS
  );

  const content = normalizeTextContent(payload?.choices?.[0]?.message?.content);
  if (!content) {
    throw new Error("No content generated.");
  }

  return content;
}

async function fetchAvailableModels(options = {}) {
  const apiKey = options.apiKey || getZaiApiKey();
  if (!apiKey) {
    throw new Error("ZAI_API_KEY is not configured.");
  }

  const payload = await fetchJson(
    `${getZaiBaseUrl()}/models`,
    {
      method: "GET",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
      },
    },
    options.timeoutMs
  );

  const discoveredModels = extractModelNames(payload);
  if (!discoveredModels.length) {
    throw new Error("No models were returned by the provider.");
  }

  return discoveredModels;
}

export {
  FALLBACK_ZAI_MODELS,
  fetchAvailableModels,
  generateText,
  getDefaultZaiModel,
  getDefaultZaiModels,
  getDefaultZaiReplyModel,
  getZaiApiKey,
  getZaiBaseUrl,
  hasZaiApiKey,
};
