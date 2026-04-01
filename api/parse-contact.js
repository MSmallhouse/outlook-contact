import Anthropic from "@anthropic-ai/sdk";

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const { text } = req.body;
  if (!text?.trim()) {
    return res.status(400).json({ error: "No text provided" });
  }

  const message = await client.messages.create({
    model: "claude-haiku-4-5-20251001",
    max_tokens: 512,
    messages: [
      {
        role: "user",
        content: `Extract contact information from the following email signature text and return ONLY a JSON object with these fields (use empty string "" for any field not found):

firstName, lastName, email, businessPhone, mobilePhone, company, jobTitle, street, city, state, zip, country, website

Text:
${text}

Return only the raw JSON object, no markdown, no explanation.`,
      },
    ],
  });

  const raw = message.content[0].type === "text" ? message.content[0].text.trim() : "";

  let contact;
  try {
    contact = JSON.parse(raw);
  } catch {
    return res.status(500).json({ error: "Failed to parse Claude response", raw });
  }

  return res.status(200).json(contact);
}
