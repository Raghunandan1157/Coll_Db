-- Phase B: AI chat history keyed by user (phone OR emp_id OR role:loc).
-- Idempotent; safe to re-run. Prefix `ai_chat_` so we don't clash with the
-- inter-employee chat tables (chat_messages, chat_threads).

DROP TABLE IF EXISTS chat_conversations CASCADE;

CREATE TABLE IF NOT EXISTS ai_chat_conversations (
  id          BIGSERIAL PRIMARY KEY,
  user_key    TEXT NOT NULL,
  title       TEXT NOT NULL DEFAULT 'New conversation',
  pinned      BOOLEAN NOT NULL DEFAULT FALSE,
  provider    TEXT,
  created_at  TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at  TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_ai_chat_conv_user
  ON ai_chat_conversations(user_key, updated_at DESC);

CREATE TABLE IF NOT EXISTS ai_chat_messages (
  id              BIGSERIAL PRIMARY KEY,
  conversation_id BIGINT NOT NULL REFERENCES ai_chat_conversations(id) ON DELETE CASCADE,
  role            TEXT NOT NULL CHECK (role IN ('user','assistant','system')),
  content         TEXT NOT NULL,
  created_at      TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_ai_chat_msg_conv
  ON ai_chat_messages(conversation_id, created_at);
