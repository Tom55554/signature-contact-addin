/*
 * taskpane.js
 *
 * This script runs inside the task pane of the Outlook add‑in. When the user
 * clicks the "Analyser depuis le mail" button, it retrieves the body of the
 * currently selected message, extracts a French phone number and a possible
 * full name from the signature, and prepopulates the form fields. When the
 * user clicks "Créer le contact", it acquires a token for Microsoft Graph
 * using Office Runtime SSO and creates a new contact using the entered
 * values. See the README in the repository root for deployment instructions.
 */

/* global Office, OfficeRuntime */

// Retrieve the body of the current mail item as HTML and plain text.
async function getItemBodyText() {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    item.body.getAsync(Office.CoercionType.Html, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        try {
          const html = res.value || "";
          const text = htmlToText(html);
          resolve({ html, text });
        } catch (e) {
          resolve({ html: res.value || "", text: res.value || "" });
        }
      } else {
        reject(res.error);
      }
    });
  });
}

// Convert simple HTML to plain text. This uses a basic approach and can be
// improved for more complex emails (e.g. handling tables, lists, etc.).
function htmlToText(html) {
  const withoutTags = html
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<[^>]+>/g, "\n");
  const decoded = withoutTags
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/\n{2,}/g, "\n")
    .trim();
  return decoded;
}

// Extract the first French phone number found in the text. This covers
// formats such as "0X XX XX XX XX", "+33 X XX XX XX XX" and numbers
// without separators (e.g. 0XXXXXXXXX).
function extractPhoneFR(text) {
  const regex = /(?:\+33\s?|\b0)(?:[1-9])(?:[\s.\-]?\d{2}){4}\b/g;
  const matches = text.match(regex);
  return matches ? normalizePhone(matches[0]) : "";
}

// Normalize a phone number: remove separators and convert +33 to a leading 0.
function normalizePhone(raw) {
  let p = raw.replace(/[.\-\s]/g, "");
  if (p.startsWith("+33")) {
    p = "0" + p.slice(3);
  }
  return p;
}

// Attempt to guess the sender's full name from the email body. It searches
// for a line in the signature following common closings (cordialement, --, etc.).
function guessFullName(text, fallbackSender) {
  const lines = text
    .split("\n")
    .map((l) => l.trim())
    .filter(Boolean);
  const sigStartIdx = lines.findIndex((l) =>
    /^(-{2,}|cordialement|bien cordialement|sincèrement|best regards|regards)$/i.test(l)
  );
  let candidateLines = lines;
  if (sigStartIdx >= 0 && sigStartIdx < lines.length - 1) {
    candidateLines = lines.slice(sigStartIdx + 1, sigStartIdx + 6);
  } else {
    candidateLines = lines.slice(-6);
  }
  for (const l of candidateLines) {
    const tokens = l.split(/\s+/).filter(Boolean);
    if (tokens.length >= 2 && tokens.length <= 4) {
      const capScore = tokens.filter((t) => /^[A-ZÀÂÄÉÈÊËÏÎÔÖÙÛÜÇ][a-zàâäéèêëïîôöùûüç'\-]+$/.test(t)).length;
      if (capScore >= 2) return l;
    }
  }
  return fallbackSender || "";
}

// Get the sender's email address and display name from the mail item.
function getSenderEmailAndName() {
  const item = Office.context.mailbox.item;
  const from = item.from || {};
  return {
    email: from.emailAddress || "",
    name: from.displayName || ""
  };
}

// Acquire a Microsoft Graph token via single sign‑on. This requires that the
// add‑in be correctly configured in Azure AD with the appropriate scopes.
async function getGraphToken() {
  return OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, forMSGraphAccess: true });
}

// Create a contact in the current user's mailbox via the Graph API.
async function createContactOnGraph(token, displayName, email, phone) {
  const body = {
    givenName: "",
    surname: "",
    displayName: displayName || email || "Nouveau contact",
    emailAddresses: email ? [{ address: email, name: displayName || email }] : [],
    businessPhones: phone ? [phone] : []
  };
  const parts = (displayName || "").split(/\s+/).filter(Boolean);
  if (parts.length >= 2) {
    body.givenName = parts[0];
    body.surname = parts.slice(1).join(" ");
  }
  const resp = await fetch("https://graph.microsoft.com/v1.0/me/contacts", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(body)
  });
  if (!resp.ok) {
    const text = await resp.text();
    throw new Error(`Graph error ${resp.status}: ${text}`);
  }
  return resp.json();
}

// Helper to set form field values.
function setVal(id, val) {
  document.getElementById(id).value = val || "";
}

// Helper to log messages to the task pane. Applies CSS classes for success
// or error states.
function log(msg, cls) {
  const el = document.getElementById("log");
  el.innerText = msg;
  el.className = `log ${cls || ""}`;
}

// Handler: scan the current message and populate the form.
async function scanAndFill() {
  log("Analyse du message en cours…");
  try {
    const { text } = await getItemBodyText();
    const { email, name } = getSenderEmailAndName();
    const phone = extractPhoneFR(text);
    const fullName = guessFullName(text, name);
    if (fullName) setVal("fullName", fullName);
    if (email) setVal("email", email);
    if (phone) setVal("phone", phone);
    log("Analyse terminée. Vérifie/édite si besoin puis clique sur “Créer le contact”.", "ok");
  } catch (e) {
    console.error(e);
    log("Échec de l’analyse : " + e.message, "ko");
  }
}

// Handler: create the contact using the entered values and Graph SSO.
async function createContact() {
  log("Création du contact…");
  try {
    const displayName = document.getElementById("fullName").value.trim();
    const email = document.getElementById("email").value.trim();
    const phone = document.getElementById("phone").value.trim();
    const token = await getGraphToken();
    const created = await createContactOnGraph(token, displayName, email, phone);
    log(
      `Contact créé ✅\nNom: ${created.displayName}\nEmail: ${created.emailAddresses?.[0]?.address || "-"}\nTel: ${created.businessPhones?.[0] || "-"}`,
      "ok"
    );
  } catch (e) {
    console.error(e);
    log(
      "Impossible de créer le contact : " + e.message +
        "\nVérifie les scopes AAD (Contacts.ReadWrite) et la configuration SSO.",
      "ko"
    );
  }
}

// Register event listeners after Office is ready. Also perform an initial scan.
Office.onReady(() => {
  document.getElementById("btnScan").addEventListener("click", scanAndFill);
  document.getElementById("btnCreate").addEventListener("click", createContact);
  scanAndFill().catch(() => {});
});