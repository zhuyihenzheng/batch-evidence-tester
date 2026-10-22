-- 生成元: docs/testcases/order_import.md（OI_001）。md を直して gen し直すこと
-- 初期データ: T_ORDER（2 行 / 削除: キー ORDER_ID が一致する行）
DELETE FROM T_ORDER WHERE [ORDER_ID] IN (N'1', N'2');
INSERT INTO T_ORDER ([ORDER_ID], [ORDER_NO], [CUSTOMER_NM], [AMOUNT], [STATUS]) VALUES
  (N'1', N'A0001', N'山田商事', N'15000', N'未処理'),
  (N'2', N'A0002', N'鈴木物産', N'8000', N'未処理');
GO
