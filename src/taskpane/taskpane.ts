import "../taskpane/taskpane.css";
import type { ParsedContact } from "../utils/parser";
import { createContact, getSignedInAccount, getAccessToken, signOut } from "../utils/graph";

// -----------------------------------------------------------------------
// View helpers
// -----------------------------------------------------------------------

type ViewId = "view-signin" | "view-loading" | "view-form";

function showView(id: ViewId): void {
  const views: ViewId[] = ["view-signin", "view-loading", "view-form"];
  for (const v of views) {
    document.getElementById(v)!.hidden = v !== id;
  }
}

function setStatus(message: string, type: "success" | "error" | ""): void {
  const bottom = document.getElementById("status-msg")!;
  const top = document.getElementById("error-msg")!;

  if (type === "error") {
    top.textContent = message;
    top.hidden = !message;
    bottom.hidden = true;
  } else {
    bottom.textContent = message;
    bottom.className = type;
    bottom.hidden = !message;
    top.hidden = true;
  }
}

function setSaveEnabled(enabled: boolean): void {
  (document.getElementById("btn-save") as HTMLButtonElement).disabled = !enabled;
}

// -----------------------------------------------------------------------
// Auth bar
// -----------------------------------------------------------------------

async function updateAuthBar(): Promise<void> {
  const account = await getSignedInAccount();
  const statusEl = document.getElementById("auth-status")!;
  const signOutBtn = document.getElementById("btn-signout") as HTMLButtonElement;

  if (account) {
    statusEl.textContent = account.username;
    signOutBtn.hidden = false;
  } else {
    statusEl.textContent = "";
    signOutBtn.hidden = true;
  }
}

// -----------------------------------------------------------------------
// Form population
// -----------------------------------------------------------------------

function populateForm(contact: ParsedContact): void {
  const fields: (keyof ParsedContact)[] = [
    "firstName", "lastName", "email",
    "businessPhone", "mobilePhone",
    "company", "jobTitle", "website",
    "street", "city", "state", "zip", "country",
  ];
  for (const field of fields) {
    const el = document.getElementById(field) as HTMLInputElement | null;
    if (el) el.value = contact[field];
  }
}

function readForm(): ParsedContact {
  const fields: (keyof ParsedContact)[] = [
    "firstName", "lastName", "email",
    "businessPhone", "mobilePhone",
    "company", "jobTitle", "website",
    "street", "city", "state", "zip", "country",
  ];
  const contact = {} as ParsedContact;
  for (const field of fields) {
    const el = document.getElementById(field) as HTMLInputElement | null;
    contact[field] = el?.value.trim() ?? "";
  }
  return contact;
}

// -----------------------------------------------------------------------
// Email reading
// -----------------------------------------------------------------------

async function getClipboardText(): Promise<string | null> {
  try {
    const timeout = new Promise<null>((resolve) => setTimeout(() => resolve(null), 1500));
    const text = await Promise.race([navigator.clipboard.readText(), timeout]);
    return text?.trim() || null;
  } catch {
    return null;
  }
}

function readEmailBody(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item!.body.getAsync(
      Office.CoercionType.Html,
      { asyncContext: null },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error(result.error?.message ?? "Failed to read email body"));
        }
      }
    );
  });
}

// -----------------------------------------------------------------------
// Main flow
// -----------------------------------------------------------------------

async function loadContact(): Promise<void> {
  showView("view-loading");
  setStatus("", "");

  try {
    const clipboardText = await getClipboardText();

    const response = await fetch("/api/parse-contact", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text: clipboardText ?? "" }),
    });

    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error((err as { error?: string }).error ?? `Server error ${response.status}`);
    }

    const contact: ParsedContact = await response.json();

    populateForm(contact);
    showView("view-form");
  } catch (err) {
    showView("view-form");
    setStatus(`Could not read email: ${(err as Error).message}`, "error");
  }
}

async function handleSignIn(): Promise<void> {
  try {
    await getAccessToken(); // triggers popup if not signed in
    await updateAuthBar();
    await loadContact();
  } catch (err) {
    setStatus(`Sign-in failed: ${(err as Error).message}`, "error");
  }
}

async function handleSignOut(): Promise<void> {
  await signOut();
  await updateAuthBar();
  showView("view-signin");
  setStatus("", "");
}

async function handleSubmit(e: Event): Promise<void> {
  e.preventDefault();
  setStatus("", "");
  setSaveEnabled(false);

  try {
    const contact = readForm();

    if (!contact.firstName && !contact.lastName) {
      setStatus("Please enter at least a first or last name.", "error");
      setSaveEnabled(true);
      return;
    }

    await createContact(contact);
    setStatus("Contact saved successfully.", "success");
    document.getElementById("btn-save")!.hidden = true;
  } catch (err) {
    setStatus(`Failed to save: ${(err as Error).message}`, "error");
    setSaveEnabled(true);
  }
}

// -----------------------------------------------------------------------
// Initialisation
// -----------------------------------------------------------------------

Office.onReady(async () => {
  // Wire up buttons
  document.getElementById("btn-signin")!.addEventListener("click", handleSignIn);
  document.getElementById("btn-signout")!.addEventListener("click", handleSignOut);
  document.getElementById("view-form")!.addEventListener("submit", handleSubmit);

  await updateAuthBar();

  const account = await getSignedInAccount();
  if (account) {
    await loadContact();
  } else {
    showView("view-signin");
  }
});
