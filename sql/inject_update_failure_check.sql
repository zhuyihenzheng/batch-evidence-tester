-- =============================================================================
-- 【方式 B】CHECK 制約 — 権限を触れない環境向け
--
--   UPLOAD_TIME に値を入れようとすると制約違反で失敗する。
--   INSERT 時は UPLOAD_TIME が NULL なので通る（batch が INSERT 時に
--   値を入れる作りなら WHERE 条件を調整すること）。
--
--   NOCHECK を付けているのは既存行を検査させないため。付けないと
--   既に値が入っている行があるだけで ALTER 自体が失敗する。
-- =============================================================================
ALTER TABLE dbo.T_ORDER WITH NOCHECK
    ADD CONSTRAINT CK_INJECT_UPLOAD_TIME CHECK (UPLOAD_TIME IS NULL);
GO
