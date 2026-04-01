import { parseContact } from "./src/utils/parser";

const testEmail = `
Hi Mark,

Nice chatting with you. I will call you when I have a client that needs your services.

Janice Hennessey, Senior Vice President
KPA Commercial, Inc.
415-676-8711

janice@kpacommercial.com
www.kpacommercial.com
CalDRE Lic. No. 01464169
`;

const result = parseContact(testEmail, "Janice Hennessey");
console.log(JSON.stringify(result, null, 2));
