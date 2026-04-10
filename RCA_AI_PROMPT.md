# DevShop RCA Extractor — AI Prompt

## Purpose
Paste this prompt into your AI tool (Gemini, ChatGPT, Claude, etc.) along with your raw incident materials. It will output a structured JSON object that you paste directly into the RCA tab of DevShop Report Studio to pre-fill the form.

---

## Prompt to Use

```
You are a ServiceNow Program Manager assistant. I am going to give you raw incident materials — which may include incident records, story exports, work notes, emails, chat logs, change records, standup notes, and any other supporting context.

Your job is to extract and organize this information into a structured JSON object that follows the exact schema below. This JSON will be loaded directly into a document generation tool, so accuracy and completeness matter.

Do not invent or assume details not present in the source materials. If a field cannot be filled from the provided content, use an empty string "" or an empty array [].

Output ONLY the raw JSON object — no explanation, no markdown code fences, no preamble.

---

SCHEMA:

{
  "incident":     "The incident number, e.g. INC0087461",
  "story":        "Related story or task number, e.g. STRY0013521",
  "change":       "Related change record, e.g. CHG0041044",
  "status":       "Current resolution status, e.g. Resolved / Closed",
  "title":        "Short description of the incident — one sentence",
  "developer":    "Name of the developer assigned to the fix",
  "reported_by":  "Name of the person who reported or discovered the issue",
  "account":      "Client or account name",
  "fix_date":     "Date the fix was applied, format YYYY-MM-DD",

  "s1": "Section 1 — Incident Overview: What happened, when it was discovered, how it was reported, and who was involved. Include any relevant context about the environment or timing.",

  "s2": "Section 2 — Root Cause Analysis: The underlying cause — configuration, logic, process, or code. Include the chain of events that led to the failure, any contributing system or platform factors, and why the issue was not caught earlier.",

  "s3": "Section 3 — Impact Assessment: Who or what was affected. Describe the scope of the disruption, any business or user impact, how many records or users were involved, and how long the issue persisted.",

  "s4": "Section 4 — Resolution Applied: What steps were taken to resolve the issue. Include what was attempted first if relevant, what the final fix was, who developed and tested it, and how it was approved and deployed.",

  "s5": "Section 5 — Contributing Factors: Process gaps, missing safeguards, environmental conditions, or communication gaps that allowed this issue to occur or to persist before it was caught.",

  "s6": [
    "Bullet item describing an immediate action taken during or right after the incident"
  ],

  "s7": [
    "Numbered step for a longer-term prevention or process improvement action"
  ],

  "s8": [
    "Any unresolved question, follow-up item, or recommended audit that remains open"
  ],

  "timeline": [
    { "time": "YYYY-MM-DD HH:MM AM/PM", "event": "Description of what happened at this timestamp" }
  ],

  "s10": "Section 10 — Document Notes: List the source materials you used (e.g. incident export, work notes, email thread, change record). Note the confidentiality classification.",

  "callouts": [
    {
      "dev": "Developer or team member name",
      "items": [
        "Please confirm [specific detail that needs verification]",
        "Please add [specific detail that is missing from the source materials]"
      ]
    }
  ]
}

---

Here are my incident materials:

[PASTE YOUR MATERIALS HERE]
```

---

## Notes

- The `callouts` array is for gaps — things the AI could not confirm from the source materials that a developer needs to verify before the document is distributed to the client. Each callout group becomes an amber box in the final document.
- `s6` is a bullet list — short action items taken immediately.
- `s7` is a numbered list — longer-term steps to prevent recurrence.
- `s8` is a list of open follow-up items that are not yet resolved.
- The timeline should use actual timestamps from the incident records where possible.
- Once you get the JSON output, switch to the RCA tab in DevShop Report Studio, paste it into the Paste JSON field on the left, and the form will pre-fill automatically.
