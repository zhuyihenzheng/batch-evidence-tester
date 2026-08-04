-- TC007 後処理: 注入したダミーデータを片付け、後続ケースの前提を壊さない
DELETE FROM T_ORDER WHERE ORDER_ID = 1 AND ORDER_NO = 'DUMMY';
GO
