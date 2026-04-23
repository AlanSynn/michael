/* global fetch, module, require */

const { ZAI_CHAT_COMPLETIONS_URL, ZAI_DEFAULT_MODEL, requireZaiApiKey } = require("./zaiConfig");

function buildZaiChatMessages(systemPrompt, userPrompt) {
  const messages = [];

  if (systemPrompt) {
    messages.push({ role: "system", content: systemPrompt });
  }

  messages.push({ role: "user", content: userPrompt });

  return messages;
}

function buildZaiChatCompletionRequest({
  systemPrompt = "",
  userPrompt,
  model = ZAI_DEFAULT_MODEL,
  temperature = 0.3,
}) {
  if (!userPrompt || !userPrompt.trim()) {
    throw new Error("Z.AI requests require a non-empty user prompt.");
  }

  return {
    model,
    messages: buildZaiChatMessages(systemPrompt, userPrompt.trim()),
    temperature,
    stream: false,
  };
}

async function parseJsonResponse(response) {
  const body = await response.text();

  if (!body) {
    return {};
  }

  try {
    return JSON.parse(body);
  } catch {
    throw new Error(`Z.AI returned a non-JSON response (${response.status}).`);
  }
}

function extractZaiMessageText(data) {
  const content = data?.choices?.[0]?.message?.content;

  if (typeof content !== "string" || !content.trim()) {
    throw new Error("Z.AI returned no message content.");
  }

  return content.trim();
}

async function executeZaiChatCompletion(request) {
  const apiKey =
    typeof request?.apiKey === "string" && request.apiKey.trim()
      ? request.apiKey.trim()
      : requireZaiApiKey();
  const payload = buildZaiChatCompletionRequest(request);
  const response = await fetch(ZAI_CHAT_COMPLETIONS_URL, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
      "Accept-Language": "en-US,en",
    },
    body: JSON.stringify(payload),
  });
  const data = await parseJsonResponse(response);
  const errorMessage = data?.error?.message || data?.message;

  if (!response.ok || errorMessage) {
    throw new Error(errorMessage || `Z.AI request failed (${response.status}).`);
  }

  return {
    text: extractZaiMessageText(data),
    data,
  };
}

module.exports = {
  buildZaiChatCompletionRequest,
  executeZaiChatCompletion,
};
