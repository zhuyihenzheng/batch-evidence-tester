-- TC008 後処理: 剥奪した権限を必ず戻す。戻さないと後続ケースが全部失敗する
REVOKE DENY UPDATE ON dbo.T_ORDER TO [batch_test_user];
GRANT UPDATE ON dbo.T_ORDER TO [batch_test_user];
GO
