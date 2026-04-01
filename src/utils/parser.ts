export interface ParsedContact {
  firstName: string;
  lastName: string;
  email: string;
  businessPhone: string;
  mobilePhone: string;
  company: string;
  jobTitle: string;
  street: string;
  city: string;
  state: string;
  zip: string;
  country: string;
  website: string;
}

// Domains to exclude from website extraction (common in email footers but not real websites)
const IGNORED_URL_DOMAINS = [
  "aka.ms", "microsoft.com", "office.com", "outlook.com",
  "linkedin.com/in", "facebook.com", "twitter.com", "x.com",
  "unsubscribe", "mailto:", "privacy", "terms",
];

/**
 * Strips quoted reply content from HTML before any other processing.
 * Removes <blockquote> blocks (Outlook's standard quoting mechanism)
 * and any div with class "gmail_quote" or similar.
 */
function stripQuotedHtml(html: string): string {
  return html
    .replace(/<blockquote[\s\S]*?<\/blockquote>/gi, "")
    .replace(/<div[^>]*class="[^"]*quote[^"]*"[\s\S]*?<\/div>/gi, "");
}

/**
 * Strips HTML tags and decodes common HTML entities from an email body.
 */
function htmlToText(html: string): string {
  return html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n");
}

/**
 * Attempts to isolate the signature block — content after a separator line
 * like "-- ", "___", or "---", or the last 15 lines as a fallback.
 */
function extractSignatureBlock(text: string): string {
  // Strip plain-text quoted replies before looking for the signature.
  // Cuts everything from "On [date] ... wrote:" or "-----Original Message-----" onward.
  const quoteHeaderPattern = /^(>.*|On .+wrote:|[-]{5}Original Message[-]{5}.*)$/im;
  const quoteMatch = text.search(quoteHeaderPattern);
  const cleanText = quoteMatch > 0 ? text.slice(0, quoteMatch) : text;

  const lines = cleanText.split("\n");

  // Look for a signature separator
  const separatorPattern = /^(\s*[-_]{2,}\s*|--\s*)$/;
  for (let i = lines.length - 1; i >= 0; i--) {
    if (separatorPattern.test(lines[i])) {
      return lines.slice(i + 1).join("\n");
    }
  }

  // Fallback: use the last 15 lines
  return lines.slice(-15).join("\n");
}

function extractEmail(text: string): string {
  const match = text.match(/[\w.+\-]+@[\w.\-]+\.[a-zA-Z]{2,}/);
  return match ? match[0] : "";
}

function extractPhone(text: string): { business: string; mobile: string } {
  // Matches formats: (555) 555-5555, 555-555-5555, +1 555 555 5555, etc.
  const phonePattern = /(\+?1[\s.\-]?)?(\(?\d{3}\)?[\s.\-]?\d{3}[\s.\-]\d{4})/g;
  const matches = [...text.matchAll(phonePattern)].map((m) => m[0].trim());

  // Heuristic: if a line containing the number also has "mobile", "cell", "m:", treat it as mobile
  const mobileKeywords = /mobile|cell|m:/i;
  let business = "";
  let mobile = "";

  for (const phone of matches) {
    const lineWithPhone = text
      .split("\n")
      .find((l) => l.includes(phone.replace(/\D/g, "").slice(-7)));
    if (lineWithPhone && mobileKeywords.test(lineWithPhone)) {
      mobile = mobile || phone;
    } else {
      business = business || phone;
    }
  }

  return { business, mobile };
}

function extractWebsite(text: string): string {
  const urlPattern = /https?:\/\/[^\s"<>()]+/gi;
  const matches = [...text.matchAll(urlPattern)].map((m) => m[0]);

  for (const url of matches) {
    const isIgnored = IGNORED_URL_DOMAINS.some((d) => url.includes(d));
    if (!isIgnored) return url;
  }
  return "";
}

/**
 * Extracts company and job title from the signature block.
 * Strategy: look for lines that follow the name and precede the phone/email,
 * typically structured as: Name → Title → Company.
 */
function extractTitleAndCompany(
  signatureText: string,
  senderName: string
): { jobTitle: string; company: string } {
  const lines = signatureText
    .split("\n")
    .map((l) => l.trim())
    .filter((l) => l.length > 0 && l.length < 80);

  // Find the line that most closely matches the sender name
  const nameParts = senderName.toLowerCase().split(" ");
  let nameLineIndex = -1;
  for (let i = 0; i < lines.length; i++) {
    const lower = lines[i].toLowerCase();
    if (nameParts.some((part) => part.length > 2 && lower.includes(part))) {
      nameLineIndex = i;
      break;
    }
  }

  // The 1-2 lines after the name are typically title and company
  const candidates = nameLineIndex >= 0
    ? lines.slice(nameLineIndex + 1, nameLineIndex + 4)
    : lines.slice(0, 4);

  // Filter out lines that look like phone numbers, emails, or URLs
  const isDataLine = (l: string) =>
    /[@+\d(]/.test(l) || /https?:\/\//.test(l) || l.includes(".com");

  const textCandidates = candidates.filter((l) => !isDataLine(l));

  return {
    jobTitle: textCandidates[0] ?? "",
    company: textCandidates[1] ?? "",
  };
}

/**
 * Attempts to extract a US/international street address from the signature.
 * Looks for a line starting with a number followed by a street name.
 */
function extractAddress(
  signatureText: string
): { street: string; city: string; state: string; zip: string; country: string } {
  const lines = signatureText.split("\n").map((l) => l.trim());

  let street = "";
  let city = "";
  let state = "";
  let zip = "";
  let country = "";

  for (let i = 0; i < lines.length; i++) {
    // Street: starts with a number and has at least one word after it
    if (!street && /^\d+\s+\w+/.test(lines[i]) && lines[i].length < 80) {
      street = lines[i];

      // Next line often has "City, ST  ZIP" or "City, State ZIP"
      const nextLine = lines[i + 1] ?? "";
      const cityStateZip = nextLine.match(
        /^([^,]+),\s*([A-Z]{2}|\w+)\s+(\d{5}(?:-\d{4})?)/
      );
      if (cityStateZip) {
        city = cityStateZip[1].trim();
        state = cityStateZip[2].trim();
        zip = cityStateZip[3].trim();
        // Line after that may be a country
        const afterCity = lines[i + 2] ?? "";
        if (afterCity && !/[@\d+]/.test(afterCity) && afterCity.length < 40) {
          country = afterCity;
        }
      }
      break;
    }
  }

  return { street, city, state, zip, country };
}

/**
 * Splits a full display name into first and last name.
 * Handles "Last, First" and "First Last" formats.
 */
function splitName(displayName: string): { firstName: string; lastName: string } {
  const name = displayName.trim();
  if (name.includes(",")) {
    const [last, first] = name.split(",").map((s) => s.trim());
    return { firstName: first ?? "", lastName: last ?? "" };
  }
  const parts = name.split(" ");
  return {
    firstName: parts[0] ?? "",
    lastName: parts.slice(1).join(" "),
  };
}

/**
 * Main entry point. Pass the raw email body (HTML or plain text) and
 * the sender's display name from Office.context.mailbox.item.from.displayName
 */
export function parseContact(rawBody: string, senderDisplayName: string): ParsedContact {
  const text = htmlToText(stripQuotedHtml(rawBody));
  const signature = extractSignatureBlock(text);

  const { firstName, lastName } = splitName(senderDisplayName);
  const email = extractEmail(signature) || extractEmail(text);
  const { business: businessPhone, mobile: mobilePhone } = extractPhone(signature);
  const website = extractWebsite(signature);
  const { jobTitle, company } = extractTitleAndCompany(signature, senderDisplayName);
  const { street, city, state, zip, country } = extractAddress(signature);

  return {
    firstName,
    lastName,
    email,
    businessPhone,
    mobilePhone,
    company,
    jobTitle,
    street,
    city,
    state,
    zip,
    country,
    website,
  };
}
