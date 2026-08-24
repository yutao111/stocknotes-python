# StockNotes 交割单分析

本地 Flask + SQLite 单体应用。当前只实现交割单导入、清洗、FIFO 处理和客观交易分析。

## 启动

双击 `start.bat`，或在当前目录执行：

```powershell
python -m pip install -r requirements.txt
python app.py
```

浏览器打开 `http://127.0.0.1:5000`。

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
