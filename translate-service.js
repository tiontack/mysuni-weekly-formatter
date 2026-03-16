"use strict";

async function translateDocumentWithAI(payload, apiKey) {
  if (!apiKey) {
    throw new Error("GEMINI_API_KEY is not configured on the server.");
  }

  const model = payload.model || "gemini-2.0-flash";
  const response = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(apiKey)}`,
    {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      systemInstruction: {
        parts: [
          {
            text:
              "You are translating a Korean corporate weekly meeting document into English. " +
              "Follow the writing guide, company glossary, and employee English-name dictionary. " +
              "Preserve the JSON structure exactly. Return valid JSON only with keys title, date, org, sections[]. " +
              "Keep dates in YYYY-MM-DD(Day) format using English weekday abbreviations Mon/Tue/Wed/Thu/Fri/Sat/Sun. " +
              "Use concise executive meeting English. Keep tables and row/column structure unchanged."
          }
        ]
      },
      contents: [
        {
          role: "user",
          parts: [
            {
              text: JSON.stringify({
                guide_items: payload.guideItems || [],
                glossary: payload.glossary || [],
                english_names: payload.names || [],
                document: payload.document || {}
              })
            }
          ]
        }
      ],
      generationConfig: {
        temperature: 0.2,
        responseMimeType: "application/json"
      }
    })
  }
  );

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(errorText || `Gemini HTTP ${response.status}`);
  }

  const result = await response.json();
  const text = extractResponseText(result);
  return parseJsonResponse(text);
}

function extractResponseText(payload) {
  const candidate = payload.candidates?.[0];
  const parts = candidate?.content?.parts || [];
  const text = parts
    .map((part) => part.text || "")
    .join("\n")
    .trim();

  if (text) {
    return text;
  }
  throw new Error("Gemini response did not contain text output.");
}

function parseJsonResponse(text) {
  const cleaned = String(text || "")
    .replace(/^```json/i, "")
    .replace(/^```/i, "")
    .replace(/```$/i, "")
    .trim();

  try {
    return JSON.parse(cleaned);
  } catch {
    const start = cleaned.indexOf("{");
    const end = cleaned.lastIndexOf("}");
    if (start >= 0 && end > start) {
      return JSON.parse(cleaned.slice(start, end + 1));
    }
    throw new Error("AI response did not contain valid JSON.");
  }
}

module.exports = {
  translateDocumentWithAI
};
