
/*
 * Copyright (c) Microsoft Corporation.
 * All rights reserved.
 * Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, DOMParser, window, fetch, URL */


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
   Load Email
   ================================ */

export async function loadEmail() {

  // clear the previous email's information
  window.phishGuardEmail = null;

  // reset the analysis results
  document.getElementById("analysis-results").textContent =
    "This email has not been analyzed yet.";

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

      console.log(
        "PhishGuard Email Data:",
        window.phishGuardEmail
      );

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

    return `
      <b>All-Link Analysis</b>
      <br><br>
      No links were found in this email.
    `;

  }

  let html = `
    <b>All-Link Analysis (${linkResults.length} links)</b>
    <br><br>
  `;

  linkResults.forEach((result, index) => {

    const findingsHTML = result.findings.length === 0
      ? "No basic URL warnings detected."
      : result.findings.map((finding) =>
          `<div>• ${escapeHTML(finding)}</div>`
        ).join("");

    html += `
      <div class="detail-item">

        <b>
          Link ${index + 1}: ${escapeHTML(result.status)}
        </b>

        <div>
          <b>Domain:</b>
          ${escapeHTML(result.domain)}
        </div>

        <div>
          <b>URL:</b>
          ${escapeHTML(result.url)}
        </div>

        <div style="margin-top: 6px;">
          ${findingsHTML}
        </div>

      </div>
    `;

  });

  return html;

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

  let color = "#16803c";

  if (risk.score >= 75) {

    color = "#c62828";

  } else if (risk.score >= 50) {

    color = "#e67e22";

  } else if (risk.score >= 25) {

    color = "#6b8e23";

  }


  /* ================================
     Build Score Explanation
     ================================ */

  const reasonsHTML = risk.reasons.length
    ? risk.reasons.map((reason) => `
        <div class="detail-item">

          <b>${escapeHTML(reason.section)}:</b>

          ${escapeHTML(reason.message)}

          <b>+${reason.points}</b>

        </div>
      `).join("")
    : "<div>No weighted indicators detected.</div>";


  // explain whether VirusTotal was included
  let note = "";

  if (vtStatus === "checked") {

    note =
      "Includes VirusTotal detections for the first link only.";

  } else if (vtStatus === "pending") {

    note =
      "Provisional score: waiting for VirusTotal.";

  } else if (vtStatus === "unavailable") {

    note =
      "Provisional score: VirusTotal report unavailable.";

  } else {

    note =
      "Local checks only; no links to check with VirusTotal.";

  }


  /* ================================
     Display Risk Score
     ================================ */

  return `
    <div style="
      padding: 16px;
      margin-bottom: 16px;
      border: 1px solid #e1e5e8;
      border-left: 5px solid ${color};
      border-radius: 6px;
      background: #ffffff;
    ">

      <div style="
        font-weight: 600;
        color: #12344d;
      ">
        PHISHING RISK SCORE
      </div>

      <div style="
        font-size: 38px;
        font-weight: 700;
        color: ${color};
        margin: 8px 0;
      ">
        ${risk.score}
        <span style="font-size: 17px;">/100</span>
      </div>

      <div style="
        font-weight: 600;
        color: ${color};
      ">
        ${risk.level}
      </div>

      <div style="
        height: 8px;
        background: #e5e7eb;
        border-radius: 4px;
        margin-top: 12px;
        overflow: hidden;
      ">

        <div style="
          width: ${risk.score}%;
          height: 100%;
          background: ${color};
        "></div>

      </div>

      <div style="margin-top: 14px;">

        <b>Why this score?</b>

        ${reasonsHTML}

      </div>

      <div style="
        margin-top: 12px;
        font-size: 11px;
        color: #6b7280;
      ">

        ${note}

        This is a rule-based risk score, not a phishing
        probability. A low score does not guarantee safety.

      </div>

    </div>
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


  // build the sender warnings
  const senderFindingsHTML =
    senderAnalysis.findings.length === 0
      ? "No basic sender warnings detected."
      : senderAnalysis.findings.map((finding) =>
          `<div>• ${escapeHTML(finding)}</div>`
        ).join("");


  const senderHTML = `
    <b>Sender Analysis</b>

    <br><br>

    <b>Domain:</b>
    <br>

    ${escapeHTML(senderAnalysis.domain)}

    <br><br>

    ${senderFindingsHTML}
  `;

  const linksHTML =
    buildLinkAnalysisHTML(linkResults);


  /* ================================
     Update Analysis Results
     ================================ */

  // this gets called whenever new analysis data is ready
  function showResults(
    virusTotalHTML,
    vtData = null,
    vtStatus = "unavailable"
  ) {

    // don't display results for an email we switched away from
    if (window.phishGuardEmail !== email) {
      return;
    }

    const risk = calculateRiskScore(
      senderAnalysis,
      linkResults,
      vtData
    );

    const scoreHTML =
      buildRiskScoreHTML(risk, vtStatus);

    results.innerHTML = `
      ${scoreHTML}

      ${senderHTML}

      <br><br>

      ${linksHTML}

      <br><br>

      <b>VirusTotal Link Analysis</b>

      <br><br>

      ${virusTotalHTML}
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

  }

}
