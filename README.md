# PhishGuard

**PhishGuard** is a Microsoft Outlook add-in designed to help users identify potentially suspicious emails by analyzing sender information and links.

The project uses JavaScript and the Microsoft Office JavaScript API to collect email information, perform basic phishing checks, and retrieve URL reputation results from VirusTotal.

## Features

- Displays the selected email's subject and sender information
- Automatically updates when switching between emails
- Extracts and displays links and attachments
- Analyzes sender addresses for unusual patterns
- Checks all extracted links for suspicious URL patterns
- Identifies recognized and unrecognized domains
- Uses VirusTotal to check the first link's reputation
- Displays analysis findings directly inside Outlook

**Note:** PhishGuard is an educational cybersecurity project. Its findings are indicators for review, not definitive proof that an email is safe or malicious.

---

## 1. Requirements

Before installing PhishGuard, make sure you have:

- Windows 10 or Windows 11
- Microsoft Outlook (classic desktop version)
- A Microsoft 365 account that supports Outlook add-ins
- Node.js and npm
- Git
- A code editor such as Visual Studio Code
- A VirusTotal API key
- An internet connection

### Install Node.js

Download and install Node.js from:

https://nodejs.org/

Choose the LTS version.

After installation, open PowerShell and verify that Node.js and npm are installed:

```powershell
node -v
npm -v
```

Both commands should display version numbers.

### Install Git

Download Git from:

https://git-scm.com/downloads

Verify the installation:

```powershell
git --version
```

---

## 2. Download PhishGuard

Open PowerShell or the terminal in Visual Studio Code.

Clone the PhishGuard repository:

```powershell
git clone YOUR_GITHUB_REPOSITORY_URL
```

Replace `YOUR_GITHUB_REPOSITORY_URL` with the actual GitHub repository address.

Navigate into the downloaded project folder:

```powershell
cd PhishGuard
```

If you want to use the development branch, run:

```powershell
git fetch origin
git switch --track origin/test
```

If the `test` branch already exists locally, use `git switch test` instead.

---

## 3. Install Project Dependencies

Inside the PhishGuard project folder, run:

```powershell
npm install
```

This installs the JavaScript packages listed in `package.json`.

The project uses packages for its Outlook add-in frontend and Express backend.

If the backend dependencies are not already included in `package.json`, install them:

```powershell
npm install express cors dotenv
```

Wait for the installation to finish before continuing.

---

## 4. Configure VirusTotal

PhishGuard uses the VirusTotal API to retrieve reputation information for links found in emails.

### Get an API key

1. Visit https://www.virustotal.com/
2. Create an account or sign in.
3. Open your account's API key section.
4. Copy your API key.

### Create the environment file

In the main PhishGuard project folder, create a file named:

```text
.env
```

Add the following:

```env
VIRUSTOTAL_API_KEY=YOUR_API_KEY_HERE
```

Replace `YOUR_API_KEY_HERE` with your own VirusTotal API key.

Save the file.

**Important:** Never upload your API key to GitHub. Make sure `.env` is included in `.gitignore`.

For example:

```gitignore
node_modules/
.env
```

The backend reads this environment variable when contacting VirusTotal.

---

## 5. Set Up the Development Certificate

PhishGuard runs its Outlook task pane using a local HTTPS development server.

Office add-ins require trusted HTTPS connections.

From the project directory, run:

```powershell
npx office-addin-dev-certs install
```

Follow the prompts to trust the development certificate.

This step may require administrator permission.

If the certificate is already installed, the tool may indicate that no changes are needed.

---

## 6. Start the PhishGuard Backend

PhishGuard uses a separate Node.js server to communicate with VirusTotal.

Open a terminal in the project folder and run:

```powershell
node server/server.js
```

If the server starts successfully, you should see:

```text
PhishGuard server running on http://localhost:3001
```

Keep this terminal open.

### Test the backend

Open a browser and navigate to:

http://localhost:3001/api/test

You should receive a response similar to:

```json
{
  "message": "PhishGuard server is running."
}
```

This confirms that the backend is responding.

It does not necessarily mean the VirusTotal API key has been validated.

---

## 7. Start the Outlook Add-in

Open a **second terminal** in the PhishGuard project folder.

Run:

```powershell
npm start
```

This starts the frontend development server and attempts to sideload PhishGuard into Outlook.

The frontend normally runs at:

https://localhost:3000

Wait for webpack to finish compiling.

A successful compilation may display:

```text
webpack compiled successfully
```

Keep this terminal open as well.

### Both servers must be running

| Server | Command | Port |
|---|---|---|
| Outlook add-in | `npm start` | 3000 |
| VirusTotal backend | `node server/server.js` | 3001 |

The local sender and link checks do not require VirusTotal, but the VirusTotal reputation check needs the backend.

---

## 8. Open PhishGuard in Outlook

1. Launch Microsoft Outlook (classic).
2. Sign in to your Microsoft 365 account.
3. Open an email.
4. Look for the PhishGuard add-in in the Outlook ribbon or add-ins menu.
5. Select the command that opens the PhishGuard task pane.

The task pane should display the email's subject, sender, links, and attachments.

If the add-in appears under an older name such as **Contoso Add-in**, check the manifest configuration and any cached add-in installation.

### Manually sideload the add-in

If PhishGuard does not appear, try:

```powershell
npx office-addin-dev-settings install manifest.xml
```

This command attempts to register the add-in using its manifest.

The exact sideloading behavior depends on the Outlook version and Microsoft 365 account configuration.

If sideloading fails with an HTTP `401` or `403` error, check that the account supports custom add-ins and that organizational policies allow them.

---

## 9. Using PhishGuard

Once the add-in is running:

1. Select an email in Outlook.
2. Wait for PhishGuard to load the email information.
3. Review the sender, links, and attachments.
4. Click **Analyze Email**.
5. Review the analysis results.

### Sender Analysis

PhishGuard checks for unusual sender-address patterns, including:

- IP-address domains
- Unusually long domains
- Multiple subdomains
- Punycode
- Unusually long usernames

### All-Link Analysis

PhishGuard checks each extracted link for:

- Recognized or unrecognized domains
- HTTP instead of HTTPS
- IP-address URLs
- Unusually long domains
- Multiple subdomains
- Punycode
- URL username or password information

An unrecognized domain is not automatically malicious. It simply means the domain is not included in PhishGuard's starter list.

### VirusTotal Analysis

PhishGuard sends the first extracted link to its local backend, which requests an existing VirusTotal URL report.

The results may include:

- Malicious detections
- Suspicious detections
- Harmless detections
- Undetected results

VirusTotal may not have an existing report for every URL.

**Privacy note:** Checking a link through VirusTotal shares that URL with an external service. Avoid submitting private or sensitive URLs without appropriate permission.

---

## 10. Troubleshooting

### `npm start` does not work

Make sure you are in the correct project folder:

```powershell
cd PhishGuard
npm install
npm start
```

If the frontend server starts but Outlook sideloading fails, try starting only the development server:

```powershell
npm run dev-server
```

This helps determine whether the problem is with webpack or the Outlook sideloading process.

### Port 3000 is already in use

Run:

```powershell
netstat -ano | findstr :3000
```

If another process is already using port 3000, close the existing development server or identify the process before restarting.

### Analysis server cannot be contacted

Make sure the backend is running:

```powershell
node server/server.js
```

Then test:

http://localhost:3001/api/test

If the backend is running but VirusTotal checks fail, inspect the backend terminal for API errors.

### VirusTotal returns an error

Check that:

- The API key is correct.
- The `.env` file is in the project root.
- The backend was restarted after editing `.env`.
- The VirusTotal API rate limit has not been exceeded.
- The computer has internet access.

### PhishGuard does not appear in Outlook

Verify that:

- Classic Outlook is being used.
- Outlook is signed in.
- The local development server is running.
- The development certificate is trusted.
- The manifest points to the correct localhost URLs.
- The Microsoft 365 account allows custom add-ins.

Restart Outlook after making changes to the manifest.

### Changes are not appearing

Save the edited files and allow webpack to recompile.

If necessary, restart the development server and reopen the add-in.

Changes to the manifest may require sideloading the add-in again.

---

## 11. Project Structure

The main files used by PhishGuard include:

```text
PhishGuard/
│
├── assets/
│   └── logo.png
│
├── server/
│   └── server.js
│
├── src/
│   └── taskpane/
│       ├── taskpane.html
│       ├── taskpane.css
│       └── taskpane.js
│
├── .env
├── .gitignore
├── manifest.xml
├── package.json
└── README.md
```

Other generated Office add-in files may also be present.

---

## 12. Technologies Used

- JavaScript
- HTML
- CSS
- Node.js
- Express.js
- Microsoft Office JavaScript API
- Microsoft Outlook
- VirusTotal API
- Webpack

---

## 13. Development Notes

PhishGuard is currently a development project and runs locally.

The frontend and backend must be started separately.

The known-domain list is intentionally limited and should not be treated as a complete list of trustworthy websites.

The VirusTotal integration currently checks only the first link, while the local URL-pattern analysis checks all extracted links.

Future development may include additional phishing indicators, improved domain reputation checks, and more detailed analysis results.

---

## License

This project includes code generated from Microsoft's Office Add-in templates, which contain MIT-licensed components. Refer to the repository's `LICENSE` file for the applicable terms.
