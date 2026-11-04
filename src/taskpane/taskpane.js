
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, DOMParser, window, fetch, URL */

// wait for Outlook before doing anything
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";

    loadEmail();

    // reload the info when another email is selected
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      loadEmail
    );

    document.getElementById("run").onclick = analyzeEmail;
  }
});

/* ================================
   Small HTML Helper
   ================================ */

// makes sure text from emails isn't treated as HTML code
function escapeHTML(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/* ================================
   Load Email
   ================================ */

export async function loadEmail() {
  // clear the old email so we don't analyze the wrong one
  window.phishGuardEmail = null;
  document.getElementById("analysis-results").textContent =
    "This email has not been analyzed yet.";

  const item = Office.context.mailbox.item;
  const output = document.getElementById("item-subject");

  if (!item) {
    output.textContent = "Select an email to view its information.";
    return;
  }

  const subject = item.subject || "";
  const senderName = item.from ? item.from.displayName || "" : "";
  const senderEmail = item.from ? item.from.emailAddress || "" : "";

  output.textContent = "Loading email information...";

  // Outlook gives us the email body through this callback
  item.body.getAsync(Office.CoercionType.Html, function (result) {
    // don't show the old email if the user switched while loading
    if (Office.context.mailbox.item !== item) return;

    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      output.textContent = "Unable to load information from this email.";
      return;
    }

    const emailBodyHTML = result.value;
    const parser = new DOMParser();
    const emailDocument = parser.parseFromString(emailBodyHTML, "text/html");

    /* ================================
       Find Links and Attachments
       ================================ */

    const links = [];

    // each <a> tag may contain a URL in its href attribute
    emailDocument.querySelectorAll("a").forEach((link) => {
      const href = link.getAttribute("href");
      if (href) links.push(href);
    });

    const attachments = item.attachments || [];

    // build the expandable lists shown in Email Information
    let linkDetails = "";
    if (links.length === 0) {
      linkDetails = "<p>No links found.</p>";
    } else {
      links.forEach((link) => {
        linkDetails += `<div class="detail-item">${escapeHTML(link)}</div>`;
      });
    }

    let attachmentDetails = "";
    if (attachments.length === 0) {
      attachmentDetails = "<p>No attachments found.</p>";
    } else {
      attachments.forEach((attachment) => {
        attachmentDetails +=
          `<div class="detail-item">${escapeHTML(attachment.name)}</div>`;
      });
    }

    /* ================================
       Show Email Information
       ================================ */

    // backticks let us insert variables into HTML with ${...}
    output.innerHTML = `
      <div class="email-field">
        <div class="field-label">SUBJECT</div>
        <div class="field-value">${escapeHTML(subject)}</div>
      </div>
      <div class="email-field">
        <div class="field-label">SENDER</div>
        <div class="sender-name">${escapeHTML(senderName)}</div>
        <div class="sender-address">${escapeHTML(senderEmail)}</div>
      </div>
      <div class="email-stats">
        <div class="stat-card">
          <div class="stat-label">LINKS</div>
          <div class="stat-count">${links.length}</div>
          <div class="stat-description">
            ${links.length === 1 ? "link detected" : "links detected"}
          </div>
          <details>
            <summary>View links</summary>
            <div class="details-content">${linkDetails}</div>
          </details>
        </div>
        <div class="stat-card">
          <div class="stat-label">ATTACHMENTS</div>
          <div class="stat-count">${attachments.length}</div>
          <div class="stat-description">
            ${attachments.length === 1 ? "attachment" : "attachments"}
          </div>
          <details>
            <summary>View attachments</summary>
            <div class="details-content">${attachmentDetails}</div>
          </details>
        </div>
      </div>
    `;

    // save the information for the Analyze Email button
    window.phishGuardEmail = {
      subject: subject,
      senderName: senderName,
      senderEmail: senderEmail,
      body: emailBodyHTML,
      links: links,
      attachments: attachments
    };

    console.log("PhishGuard Email Data:", window.phishGuardEmail);
  });
}

/* ================================
   Analyze Sender
   ================================ */

function analyzeSender(senderEmail) {
  const findings = [];
  const emailParts = senderEmail.split("@");

  if (emailParts.length !== 2) {
    return { domain: "Unknown", findings: ["Sender address could not be parsed."] };
  }

  const username = emailParts[0];
  const domain = emailParts[1].toLowerCase();

  // these are warning signs, not proof of phishing
  if (/^\d{1,3}(\.\d{1,3}){3}$/.test(domain)) {
    findings.push("Sender uses an IP address instead of a normal domain.");
  }
  if (domain.length > 50) {
    findings.push("Sender domain is unusually long.");
  }
  if (domain.split(".").length > 4) {
    findings.push("Sender address contains several subdomains.");
  }
  if (domain.includes("xn--")) {
    findings.push("Sender domain contains Punycode.");
  }
  if (username.length > 40) {
    findings.push("Sender username is unusually long.");
  }

  return { domain: domain, findings: findings };
}

/* ================================
   Known Domains
   ================================ */

// a small starter list; NOT a guarantee that a link is safe
const knownDomains = [
  "google.com", "gmail.com", "youtube.com", "microsoft.com",
  "outlook.com", "office.com", "office365.com", "live.com",
  "apple.com", "icloud.com", "amazon.com", "github.com",
  "stackoverflow.com", "wikipedia.org", "reddit.com",
  "linkedin.com", "facebook.com", "instagram.com", "paypal.com",
  "dropbox.com", "adobe.com", "zoom.us", "discord.com",
  "steampowered.com", "steamcommunity.com"
];

function isKnownDomain(domain) {
  // the dot matters: fakegoogle.com should NOT count as google.com
  return knownDomains.some((known) =>
    domain === known || domain.endsWith("." + known)
  );
}

/* ================================
   Analyze One Link
   ================================ */

function analyzeLink(link) {
  const findings = [];
  let parsedURL;

  // URL() splits a web address into protocol, hostname, etc.
  try {
    parsedURL = new URL(link);
  } catch (error) {
    return {
      url: link,
      domain: "Unknown",
      status: "Needs Review",
      findings: ["This URL could not be read."]
    };
  }

  const domain = parsedURL.hostname.toLowerCase();
  const protocol = parsedURL.protocol.toLowerCase();

  // mailto:, tel:, etc. aren't normal web links
  if (protocol !== "http:" && protocol !== "https:") {
    return {
      url: link,
      domain: domain || "Not a web address",
      status: "Not Checked",
      findings: ["Not an HTTP or HTTPS link."]
    };
  }

  const known = isKnownDomain(domain);

  if (!known) {
    findings.push("Unrecognized domain (not on our starter list).");
  }

  // basic patterns worth checking more closely
  let patternWarnings = 0;

  if (protocol === "http:") {
    findings.push("Uses HTTP instead of HTTPS.");
    patternWarnings++;
  }
  if (/^\d{1,3}(\.\d{1,3}){3}$/.test(domain)) {
    findings.push("Uses an IP address instead of a domain name.");
    patternWarnings++;
  }
  if (domain.length > 50) {
    findings.push("Domain is unusually long.");
    patternWarnings++;
  }
  if (domain.split(".").length > 4) {
    findings.push("Contains several subdomains.");
    patternWarnings++;
  }
  if (domain.includes("xn--")) {
    findings.push("Domain contains Punycode.");
    patternWarnings++;
  }
  if (parsedURL.username || parsedURL.password) {
    findings.push("URL contains username or password information.");
    patternWarnings++;
  }

  // unfamiliar doesn't automatically mean suspicious or malicious
  let status = "Recognized Domain";
  if (patternWarnings > 0) {
    status = "Suspicious Patterns";
  } else if (!known) {
    status = "Unrecognized Domain";
  }

  return { url: link, domain: domain, status: status, findings: findings };
}

/* ================================
   Analyze All Links
   ================================ */

function analyzeAllLinks(links) {
  // map() runs our checker once for each link in the list
  return links.map((link) => analyzeLink(link));
}

function buildLinkAnalysisHTML(linkResults) {
  if (linkResults.length === 0) {
    return "<b>All-Link Analysis</b><br><br>No links were found in this email.";
  }

  let html = `<b>All-Link Analysis (${linkResults.length} links)</b><br><br>`;

  linkResults.forEach((result, index) => {
    const findingsHTML = result.findings.length === 0
      ? "No basic URL warnings detected."
      : result.findings.map((finding) =>
          `<div>• ${escapeHTML(finding)}</div>`
        ).join("");

    html += `
      <div class="detail-item">
        <b>Link ${index + 1}: ${escapeHTML(result.status)}</b>
        <div><b>Domain:</b> ${escapeHTML(result.domain)}</div>
        <div><b>URL:</b> ${escapeHTML(result.url)}</div>
        <div style="margin-top: 6px;">${findingsHTML}</div>
      </div>
    `;
  });

  return html;
}

/* ================================
   Analyze Email
   ================================ */

async function analyzeEmail() {
  const results = document.getElementById("analysis-results");
  const email = window.phishGuardEmail;

  if (!email) {
    results.textContent = "Email information has not finished loading.";
    return;
  }

  const senderAnalysis = analyzeSender(email.senderEmail);
  const senderFindingsHTML = senderAnalysis.findings.length === 0
    ? "No basic sender warnings detected."
    : senderAnalysis.findings.map((finding) =>
        `<div>• ${escapeHTML(finding)}</div>`
      ).join("");

  // local checks work even if our VirusTotal server is offline
  const linkResults = analyzeAllLinks(email.links);
  const senderHTML = `
    <b>Sender Analysis</b><br><br>
    <b>Domain:</b><br>${escapeHTML(senderAnalysis.domain)}<br><br>
    ${senderFindingsHTML}
  `;
  const linksHTML = buildLinkAnalysisHTML(linkResults);

  function showResults(virusTotalHTML) {
    // don't overwrite the new email with results from the old one
    if (window.phishGuardEmail !== email) return;

    results.innerHTML = `
      ${senderHTML}<br><br>
      ${linksHTML}<br><br>
      <b>VirusTotal Link Analysis</b><br><br>
      ${virusTotalHTML}
    `;
  }

  if (email.links.length === 0) {
    showResults("No links were found in this email.");
    return;
  }

  // VirusTotal still checks only the first URL to limit API requests
  const url = email.links[0];
  showResults("Checking the first link with VirusTotal...");

  try {
    const response = await fetch("http://localhost:3001/api/check-url", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ url: url })
    });

    if (!response.ok) {
      throw new Error(`Server returned ${response.status}`);
    }

    const data = await response.json();

    if (!data.found) {
      showResults(`
        No existing VirusTotal report was found for the first link.<br><br>
        <b>Checked URL:</b><br>${escapeHTML(url)}
      `);
      return;
    }

    showResults(`
      <b>Malicious:</b> ${escapeHTML(data.malicious)}<br>
      <b>Suspicious:</b> ${escapeHTML(data.suspicious)}<br>
      <b>Harmless:</b> ${escapeHTML(data.harmless)}<br>
      <b>Undetected:</b> ${escapeHTML(data.undetected)}<br><br>
      <b>Checked URL:</b><br>${escapeHTML(url)}
    `);

    console.log("VirusTotal result:", data);
  } catch (error) {
    console.error("PhishGuard analysis error:", error);

    // keep the local results visible even if the server fails
    showResults("Could not contact the analysis server. Local checks are shown above.");
  }
}
