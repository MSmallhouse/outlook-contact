# Outlook Contact Extractor Add-in

## What This Is
An Office JS add-in for Microsoft 365 Classic Outlook (Windows desktop). When the user opens an email, a button appears in the ribbon. Clicking it opens a sidebar that parses contact info from the email body/signature, lets the user review and edit the fields, then saves the contact to Outlook via the Microsoft Graph API.

## Tech Stack
- **Office JS** — reads the current email from inside Outlook
- **TypeScript + Webpack** — compiled to static JS/HTML
- **MSAL.js** (`@azure/msal-browser`) — OAuth 2.0 sign-in for Microsoft Graph
- **Microsoft Graph API** — `POST /me/contacts` to create the contact
- **GitHub Pages** — hosts the built add-in (dev and production)

## Project Structure
```
outlook-contact/
├── CLAUDE.md
├── manifest.xml          # Sideload this into Outlook to install the add-in
├── assets/               # Required icons (16px, 32px, 80px PNG)
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html # Sidebar UI
│   │   ├── taskpane.ts   # Main logic: read email → parse → form → submit
│   │   └── taskpane.css  # Sidebar styles
│   └── utils/
│       ├── parser.ts     # Regex-based contact extraction from email body
│       └── graph.ts      # MSAL init + Graph API contact creation
├── docs/                 # Webpack output — GitHub Pages serves from here
├── package.json
├── tsconfig.json
└── webpack.config.js
```

## Prerequisites

### 1. Node & npm
```bash
node -v   # should be 18+
npm -v
```

### 2. Azure App Registration (one-time, ~5 min)
You need a Client ID so MSAL.js can authenticate with Microsoft Graph.

1. Go to https://portal.azure.com → search "App registrations" → New registration
2. Name: `OutlookContactExtractor` (anything works)
3. Supported account types: **Accounts in any organizational directory and personal Microsoft accounts**
4. Click Register
5. Copy the **Application (client) ID** — paste it into `src/utils/graph.ts` where it says `CLIENT_ID`
6. In the left sidebar: **Authentication** → Add a platform → **Single-page application**
7. Redirect URI: `https://<your-github-username>.github.io/<repo-name>/taskpane.html`
   - Also add `http://localhost:3000/taskpane.html` for local dev if needed
8. In the left sidebar: **API permissions** → Add a permission → Microsoft Graph → Delegated → search `Contacts.ReadWrite` → Add
9. No need to grant admin consent — users consent on first sign-in

## Development Workflow

### Install dependencies
```bash
npm install
```

### Build
```bash
npm run build
```
Outputs to `docs/`. GitHub Pages serves from `docs/` on the `main` branch.

### Dev loop
1. Edit code on Mac
2. `npm run build`
3. `git add docs/ && git commit -m "..." && git push`
4. GitHub Pages deploys in ~30 seconds
5. On client's Windows machine: open Outlook → right-click the task pane → **Reload**

### First-time Windows setup
1. In Outlook: **File → Manage Add-ins** (opens Outlook Web in browser)
2. Click the **+** → **Add from file**
3. Upload `manifest.xml`
4. The "Save Contact" button will appear in the ribbon when an email is open

## GitHub Pages Setup
1. Create a GitHub repo, push this project
2. Repo Settings → Pages → Source: **main branch, /docs folder**
3. Copy the Pages URL (e.g. `https://username.github.io/outlook-contact/`)
4. Update `manifest.xml`: replace the placeholder URL with your Pages URL
5. Update `src/utils/graph.ts`: set `REDIRECT_URI` to the Pages taskpane URL

## Contact Fields
The sidebar collects:
- First Name, Last Name
- Email Address
- Business Phone, Mobile Phone
- Company, Job Title
- Street, City, State, ZIP, Country
- Website

## Key Files
| File | Purpose |
|---|---|
| `manifest.xml` | Tells Outlook where the add-in lives and where to show the button |
| `src/taskpane/taskpane.ts` | Reads the email, calls the parser, renders form, calls Graph on submit |
| `src/utils/parser.ts` | Regex extraction from raw email HTML/text |
| `src/utils/graph.ts` | MSAL config + Graph API `POST /me/contacts` |

## Deployment to Client
Since the client's Windows machine is used during development, deployment is already done by the time testing is complete — the manifest is sideloaded and pointing at GitHub Pages. Just hand the machine back.

If reinstalling on a new machine: send the client `manifest.xml` and have them upload it via File → Manage Add-ins.
