# 按文件夹测试多个 batch

在项目根目录运行（Windows 可把 `python` 换成 `.venv\Scripts\python`）：

```sh
python -m autotest validate --config demo/folder_batches/config/settings.yaml
python -m autotest run --config demo/folder_batches/config/settings.yaml --offline
python -m autotest run --config demo/folder_batches/config/settings.yaml --offline --tag order
```

两个目录分别调用 `demo/fake_batch.py`，模拟两个 batch，使用独立的 sandbox。
全部运行时，结果合并到 `demo/folder_batches/output/` 的同一份 Excel。
示例不会连接 SQL Server；只验证配置切换和文件处理流程。

- `config/settings.yaml`：公共 DB、编码、报告选项，无需定义默认 exe 和 paths。
- `cases/order/settings.yaml`：订单 batch 及目录。
- `cases/invoice/settings.yaml`：发票 batch 及目录。
- 同目录的 `settings.local.yaml`：可选，本机覆盖；也支持 `setting_local.yml`。
- case 不必写 `execute.batch`；所在文件夹的 `batch:` 就是默认值。
- 截图也可分层指定：invoice 使用本文件夹的 `[processed_dir, output_dir]`；
  `ORDER_001` 用 `collect.folder_evidence.targets` 覆盖为 `[input_dir, output_dir]`。
  实行前后各截图一次；`collect.folder_evidence: []` 可关闭该 case 的文件夹截图。

迁移到自己的测试项目时，将各 batch 的设置与 case 放到相应 `cases/<batch>/`
目录，修改 exe、working_dir 和 paths 为实际路径，并删除示例的 `common_args`。
DB、环境名、Excel 输出位置仍写在根目录公共配置。配置内相对路径均以项目根目录
为基准；本示例的根目录是 `demo/folder_batches/`。
