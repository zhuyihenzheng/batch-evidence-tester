# Received TXT Checker

独立的受领 TXT 确认小程序。读取原来的布局定义 Excel，把实际 TXT 的各项数据对应到
项目名，在确认画面查看，并导出项目明细和原文 Excel。

## 启动

解压整个文件夹后，Windows 双击 `run_received_txt_checker.bat`。
需要 Python 3.6+、tkinter 和 openpyxl。Anaconda 已有这些依赖时可直接启动。
缺少 openpyxl 时，联网环境执行：

```bat
python -m pip install -r requirements-received-txt-checker.txt
```

macOS / Linux：`python3 received_txt_checker.py`，使用带 tkinter 的 Python。

## 使用

1. 点击「Excelを選択...」选择布局定义 Excel（xlsx / xlsm），选择 sheet。
2. 必要时设置「見出し行・列設定...」。默认类型 / IME / 最大桁数在 I / J / K 列。
3. 选择 TXT 的编码，默认 cp932。点击「受領TXTを選択・解析」，可多选。
   坐标可留空；TXT 完全没有坐标列时，选择「TXT項目形式」的「座標列なし」。
4. 查看项目明细、搜索或筛选异常。选择一行可查看该记录原文。
5. 点击「全件をExcel出力」。导出包含全部记录，不受画面筛选影响。

更换定义或列设置后需要重新选择 TXT 解析。程序不修改输入文件。

支持既有 raw CSV 格式：`FormID,対象有無,(FieldID,OCR値,属性,座標)...`。
坐标可为空或省略末项坐标。也支持无坐标列的三项格式：
`FormID,対象有無,(FieldID,OCR値,属性)...`，需选择「座標列なし」。
字段里的逗号、双引号、换行需符合 CSV 引用规则。FieldID 为各 FORM 展开后的连番，
日历展开为 46 项。当前不支持固定长 TXT、TSV、labeled 或单字段背面 TXT。

导出包括「項目明細」「レコード一覧」「受領原文」。
项目明细只保留：文件、记录、物理开始行、TXT 项目顺、项目名、受领 OCR 值、文字数、
最大桁数、属性、坐标、備考。FORM_ID 与对象有无仅在记录列表显示，FieldID 不输出。
備考是最后一列，后面的定义信息及人工确认列不输出。前导零与空格保留，
缺失项不会用测试值补齐。缺项、重复 ID、超长等只以黄色注意点提示，不作合否判定。
坐标为任意项，空白不提示问题。业务内容正确性由人工查看数据确认。
Excel 含受领原值，按原 TXT 的资料权限管理。

## 打包为独立 Windows EXE

在 Windows 上执行：

```bat
build_received_txt_checker.bat --install
```

生成 `dist\ReceivedTxtChecker.exe`，可单独发给没有 Python 的用户。
构建前后会执行真实 Excel → TXT 照合 → Excel 导出冒烟检查。
已有构建依赖时省略 `--install`。可通过 `CHECKER_BUILD_PYTHON` 指定构建用 Python。
也支持 `--onedir`，此时须分发 `dist\ReceivedTxtChecker\` 整个文件夹。

这个压缩包是**源码包**，不包含已编译的 Windows EXE。
