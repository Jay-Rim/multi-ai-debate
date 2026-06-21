# Multi AI Debate

Streamlit and Playwright application that lets the subscribed web versions of
Gemini, ChatGPT, and Claude debate without API billing.

The fixed workflow is:

1. Gemini explores broadly and proposes alternatives.
2. ChatGPT checks logic, evidence, operational risk, and feasibility.
3. Claude converges on a decision and execution plan.
4. ChatGPT acts as the final dissenter and challenges Claude's draft.
5. Claude produces a revised final answer.

The app stores the report locally as a Word document and can send a short
completion notification through Telegram.

## Important limitations

- This project automates consumer web interfaces. Site DOM changes can break
  selectors without warning.
- You must log in manually to each service using your own account and comply
  with each service's terms and policies.
- CAPTCHA and login confirmation may require manual interaction.
- A visible desktop session is required because the browser runs with
  `headless=False`.
- Do not commit `.env`, browser profiles, debate history, logs, or generated
  reports.

## Requirements

- Python 3.11 or newer
- Google Chrome or Microsoft Edge
- Active web subscriptions/accounts for the AI services you want to use
- Windows is the best-tested platform because long prompt entry uses the
  Windows Unicode clipboard API
- Optional: ngrok account for external access
- Optional: Telegram bot for completion notifications

## Installation

```powershell
git clone https://github.com/Jay-Rim/multi-ai-debate.git
cd multi-ai-debate
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item .env.example .env
```

Edit `.env` before running. Use a dedicated browser profile path:

```dotenv
CHROME_USER_DATA_DIR=C:/Users/your-name/ai_debate_profile
BROWSER_CHANNEL=chrome
STREAMLIT_PORT=8501
```

If you leave `BROWSER_CHANNEL` blank, install Playwright Chromium:

```powershell
python -m playwright install chromium
```

## Run

```powershell
streamlit run app.py
```

Open `http://localhost:8501`. A Chrome window opens with Gemini, ChatGPT, and
Claude tabs. Log in manually the first time. The dedicated profile preserves
those sessions for later runs.

For an ngrok tunnel in a separate process:

```powershell
python start_tunnel.py
```

The public URL is printed in the terminal and stored locally in
`ngrok_url.txt`, which is excluded from Git.

## Telegram notification

Add these optional values to `.env`:

```dotenv
TELEGRAM_BOT_TOKEN=your_bot_token
TELEGRAM_CHAT_ID=your_chat_id
```

Only a completion notice, saved filename, and available ChatGPT conversation
link are sent. The debate result and Word file are not sent.

## Data and privacy

The application writes local runtime data into:

- `history/`
- `logs/`
- `out/`
- `ai_debate_profile/` or the path configured in `.env`

All are excluded by `.gitignore`. Treat the browser profile as sensitive
because it contains authenticated sessions.

## License

MIT
