/* global __ZAI_API_KEY__, module */

const ZAI_CODING_BASE_URL = "https://api.z.ai/api/coding/paas/v4";
const ZAI_CHAT_COMPLETIONS_URL = `${ZAI_CODING_BASE_URL}/chat/completions`;
const ZAI_DEFAULT_MODEL = "glm-5-turbo";

function getZaiApiKey() {
  if (typeof __ZAI_API_KEY__ !== "string") {
    return "";
  }

  return __ZAI_API_KEY__.trim();
}

function hasZaiApiKeyConfigured() {
  return getZaiApiKey().length > 0;
}

function requireZaiApiKey() {
  const apiKey = getZaiApiKey();

  if (!apiKey) {
    throw new Error(
      "Missing ZAI_API_KEY. Set the environment variable before starting the add-in."
    );
  }

  return apiKey;
}

module.exports = {
  ZAI_CODING_BASE_URL,
  ZAI_CHAT_COMPLETIONS_URL,
  ZAI_DEFAULT_MODEL,
  getZaiApiKey,
  hasZaiApiKeyConfigured,
  requireZaiApiKey,
};
