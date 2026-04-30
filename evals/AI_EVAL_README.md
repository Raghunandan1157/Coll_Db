# AI Eval Notes

`ai_questions.jsonl` is a lightweight golden set for manual or scripted checks.

Each row contains:

- `role` and `location`: session to send with the request.
- `question`: user prompt.
- `expected_tools`: tools that should appear in streamed `tool_result` events.
- `must_include`: text or concepts expected in the final answer.
- `must_not_include`: text or concepts that indicate a regression.

Run these against `/api/ai-chat-stream` from localhost or the deployed domain with provider keys configured. For UI-only checks, use the browser and verify the behavior listed in `expected_ui`.
