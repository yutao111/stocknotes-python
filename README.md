# StockNotes 交割单分析

本地 Flask + SQLite 单体应用。当前只实现交割单导入、清洗、FIFO 处理和客观交易分析。

## Windows 免 Python 运行版

需要在一台 Windows 电脑上执行一次构建（Windows `.exe` 不能直接在 macOS 上生成）：

1. 安装 64 位 Python 3.11，仅构建电脑需要安装。
2. 双击 `build-windows.bat`。
3. 构建结果位于 `dist\StockNotes-Windows.zip`。
4. 将 ZIP 发给目标电脑，完整解压后双击 `StockNotes.exe`。目标电脑不需要安装 Python。

构建脚本会将当前 `stocknotes.db` 复制到发布包。数据库包含个人交易数据；若要发布空白版本，请在分发前删除发布目录中的 `stocknotes.db`，程序首次启动时会自动创建。

详细使用说明见 `README-WINDOWS.txt`。

## 源码启动

在当前目录执行：

```powershell
python -m pip install -r requirements.txt
python app.py
```

浏览器打开 `http://127.0.0.1:5001`。

## 数据

- 数据库默认保存在 `stocknotes.db`。
- 支持 `.xlsx`、`.xls`、`.csv`、`.txt`。
- 导入前先预览，确认后写入数据库。
- 每次确认导入后，会基于全部成交重新执行 FIFO。
- 卖出没有历史买入时保留为未匹配数据，不计入盈亏。

## 分析口径

- 买入成本：成交金额 + 佣金 + 印花税 + 其他杂费 + 其他费 + 过户费。
- 卖出净收入：成交金额 - 上述费用。
- 已实现盈亏：卖出净收入 - FIFO 匹配的买入成本。
- 胜率：盈利 FIFO 批次 / 全部已平仓 FIFO 批次。
- 当前持仓只显示成本，不估算实时市值或浮动盈亏。

运行测试：

```powershell
python -m unittest -v
```
