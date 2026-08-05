-- 方式 A の後始末。戻さないと後続の全ケースと手動実行が失敗する
REVOKE DENY UPDATE ON dbo.T_ORDER(UPLOAD_TIME) TO [batch_test_user];
GO
GRANT UPDATE ON dbo.T_ORDER(UPLOAD_TIME) TO [batch_test_user];
GO
