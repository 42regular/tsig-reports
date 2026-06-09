# Equity Research Platform

An automated AI-powered equity research system for TSIG's portfolio. Generates monthly reports covering price data, analyst consensus, financial metrics, and bull/bear theses for each holding — delivered automatically to the investment team via Google Drive on the first of every month.

## What it does

- Pulls live market data, analyst ratings, recent news, and financials for each portfolio holding using Claude AI with web search
- Structures findings into a consistent report format covering price, market cap, P/E, 52-week range, analyst consensus, financials, and bull/bear cases
- Automatically uploads a formatted Word document to a shared Google Drive folder on a monthly schedule
- Runs entirely on GitHub Actions at zero infrastructure cost

## How the monthly run works

On the 1st of every month, GitHub automatically triggers the workflow which researches each ticker sequentially, builds the report, and uploads it to Drive. The full run takes approximately 30 minutes for a 20-stock portfolio. You can also trigger it manually anytime from the Actions tab.

## Updating the portfolio

Holdings are stored as a GitHub secret (`TICKERS`) as a comma-separated list. To update positions, go to Settings → Secrets and variables → Actions → update the `TICKERS` secret with the new list.

## Stack

- **Runtime:** Node.js on GitHub Actions
- **AI:** Anthropic Claude with live web search
- **Document generation:** docx.js
- **Delivery:** Google Drive API

## Setup

See commit history for full setup steps including Google Cloud service account configuration and required GitHub secrets.
