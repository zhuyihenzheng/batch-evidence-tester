-- =============================================================================
-- 【方式 C】トリガ — 条件を細かく指定したい場合
--
--   「特定の列を更新したときだけ」「特定のレコードだけ」など、
--   失敗させる条件を厳密に決められる。前 2 方式で足りないときに使う。
-- =============================================================================
CREATE TRIGGER TRG_INJECT_UPDATE_FAIL ON dbo.T_ORDER
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    -- UPLOAD_TIME を更新したときだけ失敗させる（他の更新は通す）
    IF UPDATE(UPLOAD_TIME)
    BEGIN
        ROLLBACK TRANSACTION;
        RAISERROR('注入した障害: UPLOAD_TIME の更新に失敗しました', 16, 1);
    END
END
GO
