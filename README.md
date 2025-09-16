# Outlook AI Mail Summary App

Tools for reviewing recent Outlook conversations with the help of Microsoft Graph and OpenAI models. The project bundles a command-line assistant (`kb_mail.py`) together with a lightweight FastAPI UI (`app.py`) for summarising threads, generating replies, and optionally drafting responses directly in Outlook.

## Features
- Authenticate against Outlook.com or Microsoft 365 using the public client flow (MSAL).
- Fetch recent messages by sender or domain and generate concise summaries and reply drafts.
- Search across the mailbox with keyword queries and produce focused digests.
- Optionally create reply drafts in Outlook with the generated HTML content.
- Local FastAPI UI for non-technical users, with background launcher script.

## Prerequisites
- Python 3.10 or newer.
- An Azure App Registration configured for public client flows with delegated permissions (`Mail.Read`, `Mail.ReadWrite`, `User.Read`).
- An OpenAI API key with access to the configured `DEFAULT_MODEL`.

## Setup
1. Clone the repository and move into the project directory.
2. Copy the example environment file and fill in the secrets:
   ```bash
   cd outlook-kb-agent
   cp .env.example .env
   # edit .env with your CLIENT_ID, TENANT_ID, OPENAI_API_KEY, etc.
   ```
3. Create a virtual environment and install dependencies:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install --upgrade pip
   pip install -r requirements.txt
   ```

### Environment variables
| Variable | Description |
| --- | --- |
| `CLIENT_ID` | Azure application (public client) ID used for the Microsoft Graph login flow. |
| `TENANT_ID` | Directory scope for authentication. Use `consumers` for personal accounts, `common` for both personal and work accounts, or a specific tenant GUID. |
| `OPENAI_API_KEY` | API key used to call the OpenAI SDK. |
| `DEFAULT_MODEL` | OpenAI model identifier to use when generating summaries and replies (defaults to `gpt-4.1-mini`). |
| `TIMEZONE` | IANA timezone name for meeting proposals inserted in the generated drafts. |

The `.env` file should reside next to `kb_mail.py`. Secrets are ignored by Git; only `.env.example` is tracked for reference.

## Usage

### Command-line assistant
The CLI focuses on summarising recent conversations by sender or domain.

```bash
cd outlook-kb-agent
source .venv/bin/activate
python kb_mail.py --from-domain example.com --last 5 --days 30 --tone friendly-formal --create-draft
```

Useful switches:
- `--from-sender` / `--from-domain` (mutually exclusive) select the source of messages.
- `--last` limits the number of messages retrieved.
- `--days` filters to the last _N_ days (optional).
- `--tone` adjusts the drafting style (`brief-firm`, `friendly-formal`, `very-concise`, etc.).
- `--slot` suggests a meeting slot such as `Thu 14:00-15:00 Europe/Bucharest`.
- `--create-draft` tells the tool to create a reply draft to the newest message returned.
- `--login` pre-fills the account used for the interactive Microsoft login prompt.

### Web UI
The FastAPI UI mirrors the CLI flows with forms. Launch it with the helper script:

```bash
cd outlook-kb-agent
./run_ui_bg.sh
```

The script ensures the virtual environment exists, installs FastAPI/uvicorn if missing, starts the server on `http://127.0.0.1:8000/`, and attempts to open it in your default browser. Logs are written to `outlook-kb-agent/.logs/ui.log`.

To stop the UI, either close the terminal session or use the **Stop server** button rendered in the page, which gracefully shuts down the background process.

## Additional notes
- Authentication tokens are cached locally in `.token_cache.json` (ignored by Git).
- The OpenAI SDK is optional; if unavailable, summarisation calls will fail gracefully.
- The repository now includes a `.gitignore` to avoid committing temporary build artifacts and operating-system bundles.
