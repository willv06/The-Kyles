
/*
 * Copyright (c) Microsoft Corporation.
 * All rights reserved.
 * Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, DOMParser, window, fetch, URL */


let analysisSequence = 0;

/* ================================
   Start PhishGuard
   ================================ */

// wait for Outlook to finish loading before doing anything
Office.onReady((info) => {

  if (info.host === Office.HostType.Outlook) {

    // hide the sideload message and show PhishGuard
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";

    // load the currently selected email
    loadEmail();

    // reload everything when we select another email
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      loadEmail
    );

    // connect the Analyze Email button
    document.getElementById("run").onclick = analyzeEmail;
  }
});


/* ================================
   Small HTML Helper
   ================================ */

// keeps text from emails from being treated as HTML code
function escapeHTML(value) {

  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");

}


/* ================================
   AI Assessment Panel (UI Only)
   ================================ */

// The AI model and its API are someone else's responsibility.
// For now, this function shows a clear "not connected" state.
function buildAIAnalysisHTML(aiResult = null) {
  if (!aiResult) {
    return `
      <p class="ai-description">
        When AI is connected, this section will review the email's
        wording and context for possible phishing signs.
      </p>
      <details class="ai-details">
        <summary>What will the AI look for?</summary>
        <ul>
          <li>Urgent or threatening language</li>
          <li>Requests for passwords or personal information</li>
          <li>Impersonation and unusual instructions</li>
        </ul>
      </details>
      <p class="small-note">No AI assessment has been performed.</p>
    `;
  }

  // Expected shape: { verdict: "suspicious" | "not_suspicious" | "uncertain",
  //                   summary: "Short explanation", reasons: ["Reason 1", "Reason 2"] }
  const summary = escapeHTML(aiResult.summary || "AI analysis completed.");
  const reasons = Array.isArray(aiResult.reasons) ? aiResult.reasons : [];
  const reasonsHTML = reasons.length
    ? `<ul>${reasons.map((reason) => `<li>${escapeHTML(reason)}</li>`).join("")}</ul>`
    : "<p>No specific indicators were provided.</p>";

  return `
    <p class="ai-description">${summary}</p>
    <details class="ai-details">
      <summary>View AI indicators</summary>
      ${reasonsHTML}
    </details>
    <p class="small-note">
      AI results may be incorrect. Review the evidence before taking action.
    </p>
  `;
}

function renderAIAnalysis(result = null) {
  const slot = document.getElementById("phishguard-ai-slot");
  const badge = document.getElementById("ai-status");
  if (!slot || !badge) return;

  let label = "Not connected";
  let className = "status-badge";
  if (result) {
    const verdict = String(result.verdict || "uncertain").toLowerCase();
    if (verdict === "suspicious") {
      label = "Potential phishing";
      className += " status-badge--alert";
    } else if (verdict === "not_suspicious") {
      label = "No strong warning signs";
      className += " status-badge--clear";
    } else {
      label = "Needs review";
    }
  }

  badge.className = className;
  badge.textContent = label;
  slot.innerHTML = buildAIAnalysisHTML(result);
}

// Future integration: call this with the AI response, and the email
// object originally sent for analysis. That prevents stale results
// from appearing when someone selects a different message.
export function updateAIAnalysis(result, analyzedEmail) {
  if (!analyzedEmail || window.phishGuardEmail !== analyzedEmail) return;
  window.phishGuardAIResult = result;
  renderAIAnalysis(result);
}

/* ================================
   Load Email
   ================================ */

export async function loadEmail() {

  // Cancel stale UI updates whenever the user switches messages.
  analysisSequence++;
  window.phishGuardEmail = null;
  window.phishGuardAIResult = null;
  renderAIAnalysis(null);

  const runButton = document.getElementById("run");
  runButton.disabled = true;
  runButton.textContent = "Analyze email";

  // Reset the previous email's analysis results.
  document.getElementById("analysis-results").innerHTML =
    '<p class="empty-state">This email has not been analyzed yet.</p>';

  // grab the selected email from Outlook
  const item = Office.context.mailbox.item;

  const output = document.getElementById("item-subject");

  if (!item) {

    output.textContent =
      "Select an email to view its information.";

    return;
  }

  // collect basic email information
  const subject = item.subject || "";

  const senderName = item.from
    ? item.from.displayName || ""
    : "";

  const senderEmail = item.from
    ? item.from.emailAddress || ""
    : "";

  output.textContent = "Loading email information...";


  /* ================================
     Get Email Body
     ================================ */

  // Outlook gives us the email body through this callback
  item.body.getAsync(
    Office.CoercionType.Html,
    function (result) {

      // don't display an old email if the user switched emails
      if (Office.context.mailbox.item !== item) {
        return;
      }

      if (result.status !== Office.AsyncResultStatus.Succeeded) {

        output.textContent =
          "Unable to load information from this email.";

        return;
      }

      const emailBodyHTML = result.value;

      // turn HTML into a document JavaScript can search
      const parser = new DOMParser();

      const emailDocument = parser.parseFromString(
        emailBodyHTML,
        "text/html"
      );


      /* ================================
         Find Links
         ================================ */

      const links = [];

      // find all <a> tags in the email
      emailDocument.querySelectorAll("a").forEach((link) => {

        // href contains the actual destination URL
        const href = link.getAttribute("href");

        if (href) {
          links.push(href);
        }

      });


      /* ================================
         Find Attachments
         ================================ */

      const attachments = item.attachments || [];


      /* ================================
         Build Link List
         ================================ */

      let linkDetails = "";

      if (links.length === 0) {

        linkDetails = "<p>No links found.</p>";

      } else {

        links.forEach((link) => {

          linkDetails += `
            <div class="detail-item">
              ${escapeHTML(link)}
            </div>
          `;

        });

      }


      /* ================================
         Build Attachment List
         ================================ */

      let attachmentDetails = "";

      if (attachments.length === 0) {

        attachmentDetails = "<p>No attachments found.</p>";

      } else {

        attachments.forEach((attachment) => {

          attachmentDetails += `
            <div class="detail-item">
              ${escapeHTML(attachment.name)}
            </div>
          `;

        });

      }


      /* ================================
         Display Email Information
         ================================ */

      output.innerHTML = `
        <div class="email-field">

          <div class="field-label">SUBJECT</div>

          <div class="field-value">
            ${escapeHTML(subject)}
          </div>

        </div>

        <div class="email-field">

          <div class="field-label">SENDER</div>

          <div class="sender-name">
            ${escapeHTML(senderName)}
          </div>

          <div class="sender-address">
            ${escapeHTML(senderEmail)}
          </div>

        </div>

        <div class="email-stats">

          <div class="stat-card">

            <div class="stat-label">LINKS</div>

            <div class="stat-count">
              ${links.length}
            </div>

            <div class="stat-description">
              ${links.length === 1
                ? "link detected"
                : "links detected"}
            </div>

            <details>
              <summary>View links</summary>

              <div class="details-content">
                ${linkDetails}
              </div>
            </details>

          </div>

          <div class="stat-card">

            <div class="stat-label">ATTACHMENTS</div>

            <div class="stat-count">
              ${attachments.length}
            </div>

            <div class="stat-description">
              ${attachments.length === 1
                ? "attachment"
                : "attachments"}
            </div>

            <details>
              <summary>View attachments</summary>

              <div class="details-content">
                ${attachmentDetails}
              </div>
            </details>

          </div>

        </div>
      `;


      /* ================================
         Save Email Data
         ================================ */

      // save everything for the Analyze Email button
      window.phishGuardEmail = {

        subject: subject,
        senderName: senderName,
        senderEmail: senderEmail,
        body: emailBodyHTML,
        links: links,
        attachments: attachments

      };

      // Only enable Analyze once the email body has loaded.
      runButton.disabled = false;

    }
  );

}


/* ================================
   Analyze Sender
   ================================ */

function analyzeSender(senderEmail) {

  const findings = [];

  // split the email at the @ symbol
  const emailParts = senderEmail.split("@");

  if (emailParts.length !== 2) {

    return {
      domain: "Unknown",
      findings: [
        "Sender address could not be parsed."
      ]
    };

  }

  const username = emailParts[0];
  const domain = emailParts[1].toLowerCase();


  /* ================================
     IP Address Check
     ================================ */

  const ipAddressPattern =
    /^\d{1,3}(\.\d{1,3}){3}$/;

  if (ipAddressPattern.test(domain)) {

    findings.push(
      "Sender uses an IP address instead of a normal domain."
    );

  }


  /* ================================
     Domain Length Check
     ================================ */

  if (domain.length > 50) {

    findings.push(
      "Sender domain is unusually long."
    );

  }


  /* ================================
     Subdomain Check
     ================================ */

  if (domain.split(".").length > 4) {

    findings.push(
      "Sender address contains several subdomains."
    );

  }


  /* ================================
     Punycode Check
     ================================ */

  if (domain.includes("xn--")) {

    findings.push(
      "Sender domain contains Punycode."
    );

  }


  /* ================================
     Username Length Check
     ================================ */

  if (username.length > 40) {

    findings.push(
      "Sender username is unusually long."
    );

  }

  return {
    domain: domain,
    findings: findings
  };

}


/* ================================
   Known Domains
   ================================ */

// starter list of recognizable domains
// these are NOT guaranteed safe websites
const knownDomains = [

  "google.com",
  "gmail.com",
  "youtube.com",
  "microsoft.com",
  "outlook.com",
  "office.com",
  "office365.com",
  "live.com",
  "apple.com",
  "icloud.com",
  "amazon.com",
  "github.com",
  "stackoverflow.com",
  "wikipedia.org",
  "reddit.com",
  "linkedin.com",
  "facebook.com",
  "instagram.com",
  "paypal.com",
  "dropbox.com",
  "adobe.com",
  "zoom.us",
  "discord.com",
  "steampowered.com",
  "steamcommunity.com"

];


/* ================================
   Check Known Domain
   ================================ */

function isKnownDomain(domain) {

  // check the domain and any legitimate subdomains
  return knownDomains.some((known) => {

    return (
      domain === known ||
      domain.endsWith("." + known)
    );

  });

}


/* ================================
   Analyze One Link
   ================================ */

function analyzeLink(link) {

  const findings = [];

  let parsedURL;

  // URL() separates the address into its different parts
  try {

    parsedURL = new URL(link);

  } catch (error) {

    return {
      url: link,
      domain: "Unknown",
      status: "Needs Review",
      findings: [
        "This URL could not be read."
      ]
    };

  }

  const domain = parsedURL.hostname.toLowerCase();
  const protocol = parsedURL.protocol.toLowerCase();


  /* ================================
     Check Link Type
     ================================ */

  if (protocol !== "http:" && protocol !== "https:") {

    return {
      url: link,
      domain: domain || "Not a web address",
      status: "Not Checked",
      findings: [
        "Not an HTTP or HTTPS link."
      ]
    };

  }


  /* ================================
     Check Domain Recognition
     ================================ */

  const known = isKnownDomain(domain);

  if (!known) {

    findings.push(
      "Unrecognized domain (not on our starter list)."
    );

  }


  /* ================================
     Suspicious URL Patterns
     ================================ */

  let patternWarnings = 0;

  if (protocol === "http:") {

    findings.push(
      "Uses HTTP instead of HTTPS."
    );

    patternWarnings++;
  }

  if (/^\d{1,3}(\.\d{1,3}){3}$/.test(domain)) {

    findings.push(
      "Uses an IP address instead of a domain name."
    );

    patternWarnings++;
  }

  if (domain.length > 50) {

    findings.push(
      "Domain is unusually long."
    );

    patternWarnings++;
  }

  if (domain.split(".").length > 4) {

    findings.push(
      "Contains several subdomains."
    );

    patternWarnings++;
  }

  if (domain.includes("xn--")) {

    findings.push(
      "Domain contains Punycode."
    );

    patternWarnings++;
  }

  if (parsedURL.username || parsedURL.password) {

    findings.push(
      "URL contains username or password information."
    );

    patternWarnings++;
  }


  /* ================================
     Decide Link Status
     ================================ */

  let status = "Recognized Domain";

  if (patternWarnings > 0) {

    status = "Suspicious Patterns";

  } else if (!known) {

    status = "Unrecognized Domain";

  }

  return {
    url: link,
    domain: domain,
    status: status,
    findings: findings
  };

}


/* ================================
   Analyze All Links
   ================================ */

function analyzeAllLinks(links) {

  // map runs analyzeLink() on every URL
  return links.map((link) => analyzeLink(link));

}


/* ================================
   Build Link Analysis HTML
   ================================ */

function buildLinkAnalysisHTML(linkResults) {
  if (linkResults.length === 0) {
    return "<p>No links were found in this email.</p>";
  }

  return linkResults.map((result, index) => {
    const findingsHTML = result.findings.length === 0
      ? "No basic URL warnings detected."
      : result.findings.map((finding) =>
          `<div>${escapeHTML(finding)}</div>`
        ).join("");

    return `
      <div class="detail-item">
        <strong>Link ${index + 1}: ${escapeHTML(result.status)}</strong>
        <div><b>Domain:</b> ${escapeHTML(result.domain)}</div>
        <div><b>URL:</b> ${escapeHTML(result.url)}</div>
        <div>${findingsHTML}</div>
      </div>
    `;
  }).join("");
}


/* ================================
   Calculate Phishing Risk Score
   ================================ */

function calculateRiskScore(senderAnalysis, linkResults, vtData) {

  const reasons = [];

  let senderPoints = 0;
  let linkPoints = 0;
  let vtPoints = 0;


  // saves each warning and how many points it adds
  function addReason(section, message, points) {

    reasons.push({
      section: section,
      message: message,
      points: points
    });

  }


  /* ================================
     Sender Risk Points
     ================================ */

  const senderWeights = [

    ["Sender uses an IP address", 15],
    ["Sender domain contains Punycode", 15],
    ["Sender domain is unusually long", 5],
    ["Sender address contains several subdomains", 5],
    ["Sender username is unusually long", 5],
    ["Sender address could not be parsed", 10]

  ];

  senderAnalysis.findings.forEach((finding) => {

    const match = senderWeights.find((entry) =>
      finding.includes(entry[0])
    );

    if (match) {

      senderPoints += match[1];

      addReason(
        "Sender",
        finding,
        match[1]
      );

    }

  });


  /* ================================
     Link Risk Points
     ================================ */

  // analyze every link but only count the worst one
  // otherwise emails with many links could score too high
  let worstLink = null;

  linkResults.forEach((link, index) => {

    let points = 0;

    const linkReasons = [];

    function flag(message, weight) {

      points += weight;

      linkReasons.push({
        message: message,
        points: weight
      });

    }

    link.findings.forEach((finding) => {

      if (finding.includes("IP address")) {

        flag(finding, 15);

      } else if (finding.includes("Punycode")) {

        flag(finding, 15);

      } else if (finding.includes("HTTP instead")) {

        flag(finding, 10);

      } else if (finding.includes("username or password")) {

        flag(finding, 15);

      } else if (finding.includes("several subdomains")) {

        flag(finding, 5);

      } else if (finding.includes("unusually long")) {

        flag(finding, 5);

      } else if (finding.includes("Unrecognized domain")) {

        flag(finding, 3);

      } else if (finding.includes("could not be read")) {

        flag(finding, 10);

      }

    });

    if (!worstLink || points > worstLink.points) {

      worstLink = {
        index: index,
        points: points,
        reasons: linkReasons
      };

    }

  });

  if (worstLink) {

    linkPoints = worstLink.points;

    worstLink.reasons.forEach((reason) => {

      addReason(
        "Link " + (worstLink.index + 1),
        reason.message,
        reason.points
      );

    });

  }


  /* ================================
     VirusTotal Risk Points
     ================================ */

  // only use VirusTotal if a report was found
  if (vtData && vtData.found === true) {

    const malicious = Math.max(
      0,
      Number(vtData.malicious) || 0
    );

    const suspicious = Math.max(
      0,
      Number(vtData.suspicious) || 0
    );

    // limit points so one category can't grow forever
    const maliciousPoints = Math.min(malicious * 25, 75);
    const suspiciousPoints = Math.min(suspicious * 10, 30);

    vtPoints = maliciousPoints + suspiciousPoints;

    if (maliciousPoints > 0) {

      addReason(
        "VirusTotal",
        malicious + " engines marked the first link malicious.",
        maliciousPoints
      );

    }

    if (suspiciousPoints > 0) {

      addReason(
        "VirusTotal",
        suspicious + " engines marked the first link suspicious.",
        suspiciousPoints
      );

    }

  }


  /* ================================
     Final Risk Score
     ================================ */

  // add everything together and cap the score at 100
  const score = Math.min(
    100,
    senderPoints + linkPoints + vtPoints
  );

  let level = "Low Risk";

  if (score >= 75) {

    level = "Critical Risk";

  } else if (score >= 50) {

    level = "High Risk";

  } else if (score >= 25) {

    level = "Elevated Risk";

  }

  return {
    score: score,
    level: level,
    reasons: reasons
  };

}


/* ================================
   Build Risk Score Card
   ================================ */

function buildRiskScoreHTML(risk, vtStatus) {
  let riskClass = "";
  if (risk.score >= 75) {
    riskClass = "risk-card--critical";
  } else if (risk.score >= 50) {
    riskClass = "risk-card--high";
  } else if (risk.score >= 25) {
    riskClass = "risk-card--elevated";
  }

  const reasonsHTML = risk.reasons.length
    ? risk.reasons.map((reason) => `
        <div class="detail-item">
          <b>${escapeHTML(reason.section)}:</b>
          ${escapeHTML(reason.message)}
          <b>+${reason.points}</b>
        </div>
      `).join("")
    : "<p>No weighted indicators detected.</p>";

  let note;
  if (vtStatus === "checked") {
    note = "Includes VirusTotal detections for the first link only.";
  } else if (vtStatus === "pending") {
    note = "Provisional score: waiting for VirusTotal.";
  } else if (vtStatus === "unavailable") {
    note = "Provisional score: VirusTotal report unavailable.";
  } else {
    note = "Local checks only; no links to check with VirusTotal.";
  }

  return `
    <section class="risk-card ${riskClass}" aria-label="Rule-based phishing risk score">
      <div class="risk-label">Phishing risk score</div>
      <div class="risk-value">${risk.score}<span class="risk-max"> / 100</span></div>
      <div class="risk-level">${risk.level}</div>
      <div class="risk-track" role="progressbar" aria-label="Risk score"
           aria-valuemin="0" aria-valuemax="100" aria-valuenow="${risk.score}">
        <span class="risk-fill" style="width: ${risk.score}%;"></span>
      </div>
      <details class="risk-explanation">
        <summary>Why this score?</summary>
        <div class="risk-reasons">${reasonsHTML}</div>
      </details>
      <p class="small-note">
        ${note} This is a rule-based score, not a phishing probability.
        A low score does not guarantee safety.
      </p>
    </section>
  `;
}


/* ================================
   Analyze Email
   ================================ */

async function analyzeEmail() {

  const results =
    document.getElementById("analysis-results");

  const email = window.phishGuardEmail;

  if (!email) {

    results.textContent =
      "Email information has not finished loading.";

    return;
  }


  /* ================================
     Run Local Analysis
     ================================ */

  const senderAnalysis =
    analyzeSender(email.senderEmail);

  const linkResults =
    analyzeAllLinks(email.links);


  // Short sender summary; technical findings can be expanded below.
  const senderFindingsHTML = senderAnalysis.findings.length === 0
    ? "<p>No basic sender warnings detected.</p>"
    : senderAnalysis.findings.map((finding) =>
        `<div class="detail-item">${escapeHTML(finding)}</div>`
      ).join("");

  const senderHTML = `
    <div class="check-label">Sender domain</div>
    <p>${escapeHTML(senderAnalysis.domain)}</p>
    <div class="check-label">Findings</div>
    ${senderFindingsHTML}
  `;

  const linksHTML = buildLinkAnalysisHTML(linkResults);

  /* ================================
     Update Analysis Results
     ================================ */

  // Ignore responses from old emails or old clicks on Analyze.
  const thisAnalysis = ++analysisSequence;
  const runButton = document.getElementById("run");
  runButton.disabled = true;
  runButton.textContent = "Analyzing...";

  function finishAnalysis() {
    if (window.phishGuardEmail === email && thisAnalysis === analysisSequence) {
      runButton.disabled = false;
      runButton.textContent = "Analyze email";
    }
  }

  function showResults(virusTotalHTML, vtData = null, vtStatus = "unavailable") {
    if (window.phishGuardEmail !== email || thisAnalysis !== analysisSequence) {
      return;
    }

    const risk = calculateRiskScore(senderAnalysis, linkResults, vtData);
    const scoreHTML = buildRiskScoreHTML(risk, vtStatus);

    // The AI panel lives in taskpane.html, so this section is purely rule-based.
    results.innerHTML = `
      ${scoreHTML}
      <div class="checks-title">Technical details</div>
      <details class="check-group">
        <summary>Sender checks</summary>
        <div class="check-content">${senderHTML}</div>
      </details>
      <details class="check-group">
        <summary>Link checks (${linkResults.length})</summary>
        <div class="check-content">${linksHTML}</div>
      </details>
      <details class="check-group">
        <summary>VirusTotal (first link only)</summary>
        <div class="check-content">${virusTotalHTML}</div>
      </details>
    `;
  }


  /* ================================
     No Links Found
     ================================ */

  if (email.links.length === 0) {

    showResults(
      "No links were found in this email.",
      null,
      "no-links"
    );
    finishAnalysis();

    return;
  }


  /* ================================
     Check First Link With VirusTotal
     ================================ */

  // only the first link goes to VirusTotal for now
  const url = email.links[0];

  // display a temporary score while VirusTotal loads
  showResults(
    "Checking the first link with VirusTotal...",
    null,
    "pending"
  );


  try {

    const response = await fetch(
      "http://localhost:3001/api/check-url",
      {
        method: "POST",

        headers: {
          "Content-Type": "application/json"
        },

        body: JSON.stringify({
          url: url
        })
      }
    );


    /* ================================
       Server Error
       ================================ */

    if (!response.ok) {

      throw new Error(
        `Server returned ${response.status}`
      );

    }


    const data = await response.json();


    /* ================================
       No VirusTotal Report
       ================================ */

    if (!data.found) {

      showResults(`
        No existing VirusTotal report was found
        for the first link.

        <br><br>

        <b>Checked URL:</b>
        <br>

        ${escapeHTML(url)}
      `, null, "unavailable");

      return;
    }


    /* ================================
       Show VirusTotal Results
       ================================ */

    // recalculate the risk score using VirusTotal findings
    showResults(`
      <b>Malicious:</b>
      ${escapeHTML(data.malicious)}
      <br>

      <b>Suspicious:</b>
      ${escapeHTML(data.suspicious)}
      <br>

      <b>Harmless:</b>
      ${escapeHTML(data.harmless)}
      <br>

      <b>Undetected:</b>
      ${escapeHTML(data.undetected)}

      <br><br>

      <b>Checked URL:</b>
      <br>

      ${escapeHTML(url)}
    `, data, "checked");

    console.log(
      "VirusTotal result:",
      data
    );


  } catch (error) {

    // keep the local score if VirusTotal is unavailable
    console.error(
      "PhishGuard analysis error:",
      error
    );

    showResults(
      "Could not contact the analysis server. Local checks are shown above.",
      null,
      "unavailable"
    );

  } finally {
    finishAnalysis();
  }

}
