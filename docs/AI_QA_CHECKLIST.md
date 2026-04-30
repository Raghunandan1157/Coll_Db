# AI QA Checklist

Use this after any change to `employee.html`, `js/ai-ui.js`, or `server/index.js` AI prompts/tools.

## Access

- FO role is blocked from `/api/ai-chat`, `/api/ai-chat-stream`, `/api/voice`, and `/api/voice-stream`.
- BM, AM, DM, RM, and CEO can open the AI Assistant.
- BM cannot compare their branch with another branch.

## Kannada Voice

- Kannada reply may use Kannada words for normal text.
- All numbers must stay in English digits or English number words.
- Valid examples: `12.34 crore`, `5.6 lakh`, `45%`, `2026-05-01`.
- Invalid examples: Kannada digits, `ಕೋಟಿ`, `ಲಕ್ಷ`, Kannada number words for dashboard numbers.

## Tool Reasoning

- Relative dates call `resolve_date_range` before data tools.
- Named branches call `find_branch` or are canonicalized by the server before performance queries.
- Named employees use `find_employee`; ambiguous results ask the user to choose.
- Daily plan, FTOD, DPD, KYC, and NPA closure questions use `daily_reports_query`.
- Month-over-month and week-over-week questions use `period_compare`.
- Branch health questions use `branch_summary`.
- Drilldown questions use `collection_drilldown`.

## UI

- AI Assistant appears once in the desktop sidebar.
- Provider chooser disables unavailable providers.
- Chat stream shows thinking, tool, answer, and done states.
- Markdown tables render cleanly in chat.
- Copy on an AI table copies a PNG image where supported; fallback downloads an image or copies TSV only if image copy cannot run.
- Voice/tool-result table copy also prefers PNG image.
