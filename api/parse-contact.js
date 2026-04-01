const Anthropic = require("@anthropic-ai/sdk");

const client = new Anthropic.default({ apiKey: process.env.ANTHROPIC_API_KEY });

module.exports = async function handler(req, res) {
  // Allow the Outlook task pane to call this function
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const { text } = req.body ?? {};
  if (!text?.trim()) return res.status(400).json({ error: "No text provided" });

  try {
    const message = await client.messages.create({
      model: "claude-haiku-4-5-20251001",
      max_tokens: 512,
      messages: [
        {
          role: "user",
          content: `Extract contact information from the following email signature text and return ONLY a JSON object with these exact fields (use empty string "" for any field not found):

firstName, lastName, email, businessPhone, mobilePhone, company, jobTitle, street, city, state, zip, country, website

Text:
${text}

Return only the raw JSON object, no markdown, no explanation.`,
        },
      ],
    });

    const raw = message.content[0].type === "text" ? message.content[0].text.trim() : "";
    const cleaned = raw.replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/, "").trim();
    const contact = JSON.parse(cleaned);
    return res.status(200).json(contact);
  } catch (err) {
    return res.status(500).json({ error: err.message ?? "Unknown error" });
  }
};
