# GPT‑4o Chat Box for Google Sheets

A 40‑line Apps Script that turns any Google Sheet into a lightweight ChatGPT
client—perfect for teammates without a ChatGPT account.

## Features
* **Single‑sheet UI** – three named cells (`Input`, `Reply`, `Run`) keep it clean.
* **BYO API Key** – paste your OpenAI key in `BYO_Key`; no shared billing.
* **Dropdown or drawing button** – assign `runChat()` to a drawing for 1‑click send.
* **Conversation memory** – last 20 turns stored in `DocumentProperties`.
* **Reset menu** – clear history with one click.

## Quick Start
1. `File ▸ Make a copy` of the template Sheet in `/template`.
2. Open **Extensions ▸ Apps Script** and paste `ChatGPT.gs`.
3. In **Script properties**, add  
   `OPENAI_KEY = sk-…`  *(optional if users supply their own keys)*  
4. Back in the Sheet:  
   * Define named ranges: `BYO_Key`, `Input`, `Reply`, `Run`.  
   * Insert a Drawing, label it `▶ Run`, right‑click ▸ **Assign script** ▸ `runChat`.  
5. Type a prompt → click **▶ Run** → get your reply.

## Deploying Updates
Push edits to `ChatGPT.gs`, then in the Sheet open **Apps Script ▸ Deploy ▸ Test deployments** to refresh.

## License
MIT – do what you want, credit appreciated.
