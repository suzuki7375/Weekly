# Weekly PPT 合併器

這是一個桌面小工具（Tkinter），可協助你：

1. 選擇「範例文件」(pptx)
2. 選擇「整合資料」資料夾（內含多份人員週報 pptx）
3. 按下「合併並撰寫」後：
   - 依檔名排序，將整合資料資料夾內所有 pptx 投影片附加到範例文件後方
   - 在範例文件第 2 頁（page 2）設定標題為 `System team weekly Status`
   - 讀取每份週報內容，為每位同仁整理最多 4 點重點，寫入第 2 頁

## 安裝

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 執行

```bash
python app.py
```

## 備註

- 範例文件至少要有 2 頁，程式才可在第 2 頁寫入總結。
- 人名預設以檔名（不含副檔名）判定。
