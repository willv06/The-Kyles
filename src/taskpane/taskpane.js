/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, DOMParser, window, fetch */

// waits until the Office/Outlook stuff is ready 
Office.onReady((info) => {

  // makes sure this is actually running inside Outlook
  if (info.host === Office.HostType.Outlook) {

    // hide the message that says the add-in needs to be opened in Outlook
    document.getElementById("sideload-msg").style.display = "none";

    // show the actual PhishGuard interface
    document.getElementById("app-body").style.display = "block";

    // load whatever email is currently selected
    loadEmail();

    // Outlook runs loadEmail again whenever we click on a different email so PhishGuard updates without needing to reopen it
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      loadEmail
    );

    // when the Analyze Email button is clicked, run analyzeEmail()
    document.getElementById("run").onclick = analyzeEmail;
  }
});


/* ================================
   Load Email
   ================================ */

export async function loadEmail() {

  // clear the old email data so we dont accidentally analyze the last email
  window.phishGuardEmail = null;

  // also reset whatever analysis results were showing before
  document.getElementById("analysis-results").innerHTML =
    "This email has not been analyzed yet.";

  // "item" is the email that is currently open/selected in Outlook
  const item = Office.context.mailbox.item;

  // cant proceed w/out email selected
  if (!item) {

    document.getElementById("item-subject").innerHTML =
      "Select an email to view its information.";

    // return stops the function
    return;
  }

  // grab some basic info directly from the Outlook email
  const subject = item.subject;
  const senderName = item.from.displayName;
  const senderEmail = item.from.emailAddress;

  // save the HTML element where we display all the email info
  
  const output = document.getElementById("item-subject");

  output.innerHTML = "Loading email information...";


  // get the actual email body from Outlook
  
  item.body.getAsync(
    Office.CoercionType.Html,
    function (result) {

      // only continue if Outlook successfully gave us the email body
      if (result.status === Office.AsyncResultStatus.Succeeded) {

        const emailBodyHTML = result.value;


        // turn the email HTML into a document JS can search through
        // this lets us do stuff like find all the <a> tags
        const parser = new DOMParser();

        const emailDocument = parser.parseFromString(
          emailBodyHTML,
          "text/html"
        );


        /* ================================
           Find Links
           ================================ */

        // querySelectorAll("a") finds every <a> tag in the email
        // <a> tags are what HTML uses for links
        const linkElements =
          emailDocument.querySelectorAll("a");

        // [] makes an empty array
        // we'll fill this with every URL we find
        const links = [];

        // forEach goes through each link we found one at a time
        linkElements.forEach((link) => {

          // href contains the actual URL the link points to
          const href = link.getAttribute("href");

          // some <a> tags might not actually have an href
          // so only save it if one exists
          if (href) {

            // push adds the URL to the end of our links array
            links.push(href);
          }

        });


        /* ================================
           Find Attachments
           ================================ */

        // Outlook already gives us a list of attachments
        // || [] basically means "use an empty list if there arent any"
        const attachments = item.attachments || [];


        /* ================================
           Build Link List
           ================================ */

        // this will eventually contain the HTML for our list of links
        let linkDetails = "";

        if (links.length === 0) {

          linkDetails = "<p>No links found.</p>";

        } else {

          // go through each URL we found and make a little div for it
          links.forEach((link) => {

            // += adds more HTML instead of replacing what is already there
            // ${link} puts the value of the variable into the HTML
            linkDetails += `
              <div class="detail-item">
                ${link}
              </div>
            `;

          });

        }


        /* ================================
           Build Attachment List
           ================================ */

        let attachmentDetails = "";

        if (attachments.length === 0) {

          attachmentDetails =
            "<p>No attachments found.</p>";

        } else {

          // same idea as the links above
          // just using the attachment's filename instead
          attachments.forEach((attachment) => {

            attachmentDetails += `
              <div class="detail-item">
                ${attachment.name}
              </div>
            `;

          });

        }


        /* ================================
           Show Email Information
           ================================ */

        // innerHTML lets us build the information card with HTML
        // the backticks (`) let us make a multi-line string
        // ${variable} inserts JS values into it
        output.innerHTML = `

          <div class="email-field">

            <div class="field-label">
              SUBJECT
            </div>

            <div class="field-value">
              ${subject}
            </div>

          </div>


          <div class="email-field">

            <div class="field-label">
              SENDER
            </div>

            <div class="sender-name">
              ${senderName}
            </div>

            <div class="sender-address">
              ${senderEmail}
            </div>

          </div>


          <div class="email-stats">

            <div class="stat-card">

              <div class="stat-label">
                LINKS
              </div>

              <div class="stat-count">
                ${links.length}
              </div>

              <div class="stat-description">
                ${links.length === 1
                  ? "link detected"
                  : "links detected"}
              </div>

              <details>

                <summary>
                  View links
                </summary>

                <div class="details-content">
                  ${linkDetails}
                </div>

              </details>

            </div>


            <div class="stat-card">

              <div class="stat-label">
                ATTACHMENTS
              </div>

              <div class="stat-count">
                ${attachments.length}
              </div>

              <div class="stat-description">
                ${attachments.length === 1
                  ? "attachment"
                  : "attachments"}
              </div>

              <details>

                <summary>
                  View attachments
                </summary>

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

        // save everything we collected in one place
        // analyzeEmail() can grab this later when the button is clicked
        window.phishGuardEmail = {

          subject: subject,
          senderName: senderName,
          senderEmail: senderEmail,
          body: emailBodyHTML,
          links: links,
          attachments: attachments

        };

        // useful for testing/debugging
        // lets us see all the collected email data in the console
        console.log(
          "PhishGuard Email Data:",
          window.phishGuardEmail
        );

      } else {

        // if Outlook couldnt give us the email body
        output.innerHTML =
          "Unable to load information from this email.";

      }

    }
  );
}


/* ================================
   Analyze Sender
   ================================ */

function analyzeSender(senderEmail) {

  // this array will hold any weird/suspicious things we notice
  const findings = [];


  // split breaks a string into pieces wherever it finds "@"
  // example:
  // bob@example.com becomes ["bob", "example.com"]
  const emailParts = senderEmail.split("@");


  // a normal email should give us exactly 2 pieces
  if (emailParts.length !== 2) {

    return {
      domain: "Unknown",
      findings: [
        "Sender address could not be parsed."
      ]
    };

  }


  // [0] means the first item in the array
  // [1] means the second item
  const username = emailParts[0];

  // convert domain to lowercase so capitalization doesnt affect our checks
  const domain = emailParts[1].toLowerCase();


  /* ================================
     Check For IP Address
     ================================ */

  // this pattern looks for something shaped like an IPv4 address
  // example: user@192.168.1.1
  const ipAddressPattern =
    /^\d{1,3}(\.\d{1,3}){3}$/;

  // .test() checks whether the domain matches the pattern above
  if (ipAddressPattern.test(domain)) {

    findings.push(
      "Sender uses an IP address instead of a normal domain."
    );

  }


  /* ================================
     Check Domain Length
     ================================ */

  // very long domains arent automatically bad
  // but they can be worth pointing out
  if (domain.length > 50) {

    findings.push(
      "Sender domain is unusually long."
    );

  }


  /* ================================
     Check Subdomains
     ================================ */

  // splitting on "." gives us each section of the domain
  // mail.example.com would become:
  // ["mail", "example", "com"]
  const domainParts = domain.split(".");

  if (domainParts.length > 4) {

    findings.push(
      "Sender address contains several subdomains."
    );

  }


  /* ================================
     Check For Punycode
     ================================ */

  // xn-- means the domain is using Punycode
  // Punycode can be completely legitimate
  // but phishing sites can also use it for lookalike domains
  if (domain.includes("xn--")) {

    findings.push(
      "Sender domain contains Punycode."
    );

  }


  /* ================================
     Check Username Length
     ================================ */

  // again, this doesnt automatically mean phishing
  // we're just pointing out something unusual
  if (username.length > 40) {

    findings.push(
      "Sender username is unusually long."
    );

  }


  // send the domain and everything we found back to analyzeEmail()
  return {
    domain: domain,
    findings: findings
  };

}


/* ================================
   Analyze Email
   ================================ */

// async is needed because later we use "await" while contacting our server
async function analyzeEmail() {

  // where we will display the analysis results
  const results =
    document.getElementById("analysis-results");


  // make sure loadEmail() actually finished first
  if (!window.phishGuardEmail) {

    results.innerHTML =
      "Email information has not finished loading.";

    return;
  }


  // grab the links we saved earlier
  const links = window.phishGuardEmail.links;


  // run the sender email through our sender checks
  const senderAnalysis = analyzeSender(
    window.phishGuardEmail.senderEmail
  );


  // build the text that will show our sender findings
  let senderFindingsHTML = "";

  if (senderAnalysis.findings.length === 0) {

    senderFindingsHTML = `
      No basic sender warnings detected.
    `;

  } else {

    // make a new line for every sender warning we found
    senderAnalysis.findings.forEach((finding) => {

      senderFindingsHTML += `
        <div>
          • ${finding}
        </div>
      `;

    });

  }


  /* ================================
     If There Are No Links
     ================================ */

  // sender analysis can still work even if there arent any URLs
  if (links.length === 0) {

    results.innerHTML = `
      <b>Sender Analysis</b>

      <br><br>

      <b>Domain:</b>
      <br>
      ${senderAnalysis.domain}

      <br><br>

      ${senderFindingsHTML}

      <br><br>

      <b>VirusTotal Link Analysis</b>

      <br><br>

      No links were found in this email.
    `;

    // no URL means theres nothing for VirusTotal to check
    // so we stop the function here
    return;
  }


  /* ================================
     Check First Link
     ================================ */

  // [0] means grab the first URL in the links array
  // for now PhishGuard only checks the first link
  const url = links[0];

  results.innerHTML = `
    <b>Checking link...</b>

    <br><br>

    Contacting VirusTotal.
  `;


  // try means "attempt this code"
  // if something fails, the catch section at the bottom handles it
  try {

    // fetch sends a request to our PhishGuard backend server
    // await means wait for the server to answer before moving on
    const response = await fetch(
      "http://localhost:3001/api/check-url",
      {
        // POST lets us send information to the server
        method: "POST",

        // tell the server that the information we're sending is JSON
        headers: {
          "Content-Type": "application/json"
        },

        // JSON.stringify turns our JS object into JSON text
        // this sends the URL we want checked
        body: JSON.stringify({
          url: url
        })
      }
    );


    /* ================================
       Server Error
       ================================ */

    // response.ok is false for HTTP errors like 404 or 500
    if (!response.ok) {

      // throw sends us down to the catch(error) section
      throw new Error(
        `Server returned ${response.status}`
      );

    }


    // turn the JSON response from our server back into a JS object
    const data = await response.json();


    /* ================================
       VirusTotal Has No Report
       ================================ */

    // our backend sets found to false if VirusTotal doesnt know the URL
    if (!data.found) {

      results.innerHTML = `
        <b>Sender Analysis</b>

        <br><br>

        <b>Domain:</b>
        <br>
        ${senderAnalysis.domain}

        <br><br>

        ${senderFindingsHTML}

        <br><br>

        <b>VirusTotal Link Analysis</b>

        <br><br>

        No existing VirusTotal report was found
        for the first link in this email.

        <br><br>

        <b>Checked URL:</b>
        <br>
        ${url}
      `;

      return;
    }


    /* ================================
       Show VirusTotal Results
       ================================ */

    // if we got this far, VirusTotal had a report
    // data contains the numbers our backend got from VirusTotal
    results.innerHTML = `
      <b>Sender Analysis</b>

      <br><br>

      <b>Domain:</b>
      <br>
      ${senderAnalysis.domain}

      <br><br>

      ${senderFindingsHTML}

      <br><br>

      <b>VirusTotal Link Analysis</b>

      <br><br>

      <b>Malicious:</b> ${data.malicious}
      <br>

      <b>Suspicious:</b> ${data.suspicious}
      <br>

      <b>Harmless:</b> ${data.harmless}
      <br>

      <b>Undetected:</b> ${data.undetected}

      <br><br>

      <b>Checked URL:</b>
      <br>
      ${url}
    `;


    // put the full VirusTotal response in the console for debugging
    console.log(
      "VirusTotal result:",
      data
    );


  } catch (error) {

    // if fetch fails or we threw an error above,
    // print the real error in the console so we can troubleshoot it
    console.error(
      "PhishGuard analysis error:",
      error
    );

    // give the user a simpler error message in the actual add-in
    results.innerHTML = `
      <b>Analysis Error</b>

      <br><br>

      PhishGuard was unable to contact
      the analysis server.
    `;

  }

}