# DevShop-Status-Report-Generator
# DevShop Report Studio v2.0

A native Windows desktop app for ServiceNow Program Managers.
Generates polished Status Reports, Call Scripts, and Root Cause 
Analysis documents from structured JSON — no browser, no server, 
no installation required for end users.

## How It Works

1. Run your standup and gather raw inputs
2. Paste materials into your AI tool with the extraction prompt
3. Drop or paste the JSON output into DevShop Report Studio
4. Review and edit in the app
5. Click Generate — Word document saved to Documents/DevShop Studio/output/

## Requirements (Build Machine Only)

- Node.js v16 or higher
- Windows 10/11 64-bit

End users need only the .exe and Microsoft Word.

## Build

Double-click BUILD.bat — it installs dependencies and produces:
dist/DevShop_Report_Studio.exe

## Document Modes

| Mode | Trigger | Output |
|---|---|---|
| Status Report | JSON has executive_summary field | Weekly status .docx |
| Call Script | JSON has topics[] array | Structured call agenda .docx |
| RCA | JSON has incident field | Root cause analysis .docx |

## AI Extraction

Use any capable AI tool (Claude, Gemini, ChatGPT) with the prompts in:
- RCA_AI_PROMPT.md (for RCA mode)
- Your Gemini Gem or equivalent setup for Status Report and Call Script

## Docs

Full technical reference in /docs.

## Version

v2.0 — April 2026
