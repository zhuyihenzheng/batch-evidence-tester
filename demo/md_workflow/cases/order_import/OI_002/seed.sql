-- 生成元: docs/testcases/order_import.md（OI_002）。md を直して gen し直すこと
-- 初期データ: T_ORDER（2 行 / 削除: キー ORDER_ID が一致する行）
DELETE FROM T_ORDER WHERE [ORDER_ID] IN (N'3', N'4');
INSERT INTO T_ORDER ([ORDER_ID], [ORDER_NO], [CUSTOMER_NM], [AMOUNT], [STATUS]) VALUES
  (N'3', N'A0003', N'佐藤工業', N'12000', N'未処理'),
  (N'4', N'A0004', N'田中電機', N'-500', N'未処理');
GO
