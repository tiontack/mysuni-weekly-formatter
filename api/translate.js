"use strict";

const { translateDocumentWithAI } = require("../translate-service");

module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  try {
    const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body || {};
    const document = await translateDocumentWithAI(body, process.env.GEMINI_API_KEY);
    res.status(200).json({ document });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message || "Translation failed" });
  }
};
