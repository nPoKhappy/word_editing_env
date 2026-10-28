# OfficeCLI 聊天室復現與可行性驗證

日期：2026-10-28（Asia/Taipei）

**結論：OfficeCLI 可以取代本次一般 DOCX 正文、表格與圖片的 Python 文件操作；完整的 Word／EndNote 工作流程仍需保留原生介面。** 本次沒有以「複製最後一版」代替驗證：從 v6 的空白第四章出發，用 OfficeCLI 分別重建 v7–v15 九個階段，再以原生 Word 建立標號、交互參照和最終分頁。

來源聊天室：[擬定第四章結果與討論雛形](https://chatgpt.com/s/cx_6ae1bcff28c081919e22559d0f9c375d)。直接網頁擷取失敗後，已用瀏覽器讀到完整可見對話。原有中間稿作為內容／格式規格，並非復現產出。這驗證編輯能力與保留性；不代表重新執行模型訓練、Aspen 模擬或重新驗證研究結論。

## 明天先看這些檔案

所有產出都在本專案 `outputs/chat-replay-20261028/`：

- `delivery/JTICE_OfficeCLI_聊天室復現完成版.docx`：主交付文件。
- `render/delivery/JTICE_OfficeCLI_聊天室復現完成版.pdf`：Word 原生匯出的 61 頁 PDF。
- `delivery/JTICE_驗證案例原始RMSE_補充資料.csv`：原始 RMSE 補充資料。
- `stages/`：九個歷史階段的 OfficeCLI 重建文件；保留失敗嘗試與批次紀錄以供追查。
- `logs/`：內容、數值、字級、原生欄位、MCP、測試與渲染盤點。
- `plots/`：本次新生成的圖檔、向量 PDF、數據座標檢查與來源雜湊。
- `specs/`：從原稿唯讀擷取的結構化重建規格與圖片。

最終交付版以 OfficeCLI 插入 10 張本次重繪圖，其他原始圖沿用原有素材。前面的超參數選擇保留四種 ρ；後面的比較保留 No-Gain、Gain（0.05）、Gain + SS（0.01）及 Difficult-GradNorm（ρ = 0.01）。

## 逐階段復現

| 階段 | 操作 | OfficeCLI 成功操作數 | 檢驗 |
|---|---|---:|---|
| v7 | 第四章正文、ANN 與步階、選參、KCI、動態預測，新增圖表 | 1,384 | 通過 |
| v8 | 合併 ANN 圖、無標題圖片、五條 RMSE 曲線 | 1,390 | 通過 |
| v9 | 逐案例原始 RMSE 表與最低值粗體 | 3,486 | 通過 |
| v10 | 逐案例表格的整理版 | 3,399 | 通過 |
| v11 | 與 No-Gain 比較的紅綠色，移除平均 RMSE 表 | 3,426 | 通過 |
| v12 | RMSE 加相對變化率的雙行儲存格 | 3,693 | 通過 |
| v13 | 四張單行百分比表、選定 ρ 灰底 | 3,211 | 通過 |
| v14 | 分拆參考方向／大小圖、四個 gain pair 詳細討論 | 3,741 | 通過 |
| v15 | 後續 KCI、RMSE 圖表恢復四模型比較 | 3,405 | 通過 |
| **合計** | 不含最後重新生成圖的 10 次替換 | **27,135** | **九階段通過** |

各階段都從同一 v6 起始稿以實際 OfficeCLI 批次重建，記錄讀回與結構驗證；這不是逐字重演原先的 Python 程式。內容對照使用原生 REF 的實際書籤目標，另記錄快取差異；空白段落的排版由最終 Word 渲染檢查。

## 具體驗證結果

- 九階段 DOCX 結構驗證皆為零錯誤；章節文字、表格、圖片位元內容、公式及前後章文字的對照通過。對照階段使用原始圖片；交付版換入本次新繪的十張圖另有雜湊驗證。
- 原始 CSV 獨立重算：每個表格階段 102 列；後期每版 120 個百分比及紅綠色，分表版 30 個選定 ρ 灰底儲存格；最低值皆依未四捨五入的資料判定。
- 最終四模型表：288 個 RMSE 值、120 個百分比、408 個最低值粗體判斷、120 個基準顏色判斷全部通過。
- 26 個 genuine EndNote ADDIN 欄位的完整 code 與 cached result 與 v6 相同；21 個數學物件保留。
- 第四章正文、表格、圖說共 177 處 H₂S／SO₂ 的真正原生下標檢查通過。
- 所有階段新增或重製表格都經原生 Word 檢查普通與複雜字型的 12 pt；沒有為塞入表格而縮為 10 pt。
- 每階段均以選擇性原生更新完成欄位和分頁。此期刊稿原本沒有主目錄、圖表目錄，因此目錄字級為不適用；沒有憑空插入論文式目錄。
- 完整 OfficeCLI 整合測試 23 項通過；原有 EndNote 單元／回歸測試 58 項通過。
- 四個專案 MCP 均以本專案設定重新啟動、列出工具並成功實際呼叫一個唯讀工具；既有聊天中的服務需重新載入才會套用新路徑。
- 新圖直接核對 170 條五模型 RMSE 曲線、96 條四模型 RMSE 曲線、ANN 預測點及八條操作軌跡；三張四模型 KCI 圖核對 224 個 CSV 值。空的負方向群組維持 N/A。
- 最終 Word PDF 共 61 頁，已逐頁檢查整體版面；重點表格和圖另以單頁檢查。頁面清單見 `logs/render-inventory.json`。
- 90 個已盤點原始來源檔案的雜湊保持不變。OfficeCLI 的通用 issues 掃描另有 154 個「未設定首行縮排」和 39 個「空白段落」提示，已對照原稿的區塊式排版與渲染檢查；沒有據此擅自重設全稿縮排。這些提示不等同於結構驗證錯誤。

## 修正的工具問題

1. `edit_office.ps1` 原本仍指定舊專案 TEMP；改由目前 checkout 推導 `.tmp`。
2. 專案 AGENTS、相關 skills 及 MCP 設定的舊執行路徑改至本專案；EndNote 套件、狀態與搜尋快照也獨立複製，實際文獻庫不改動。
3. 圖片替換使用 OfficeCLI 回傳的 drawing run 路徑；不假設 `/picture[1]` 子路徑存在。失敗批次已驗證為整批回復。
4. OfficeCLI run 的 `text` 不解讀段落的 `\v` 跳脫字串；雙行儲存格改為明確新增 `break`、`type=line`。
5. Word `InsertCaption` 以保留佔位段落後精確移除的方式處理，避免產生多餘空白段落；書籤依原生欄位範圍建立，不以顯示字串長度當 Word Range 長度。
6. 原生 caption 巡覽改為一次讀取 SEQ 欄位，避免反覆跨程序掃描數千個表格段落。
7. 表格驗證片段若以負號開頭，使用 `--table-text=...` 傳遞，避免 argparse 將負百分比誤當選項。
8. pip 與 matplotlib 的後續快取／設定路徑明確放在本專案 `.tmp`。
9. 長 PDF 的連續 PNG 渲染曾在第 58 頁漏畫部分文字；PDF 與獨立單頁轉圖均完整。`render_docx.py` 已改為有期限的逐頁獨立程序，並明確保留 bitmap 到儲存完成，再重新渲染全文件。

## 邊界與尚未解決的原稿問題

| 工作 | 結果 |
|---|---|
| 封閉 DOCX 正文、run、可編輯表格、數字格式、圖片 | OfficeCLI 可替代一般 Python 文件編輯 |
| 原生 InsertCaption、交互參照與真正可更新的書籤範圍 | 保留 Word 原生 adapter |
| EndNote／CWYW 欄位、文獻庫操作 | 保留既有原生工具，不能以手打欄位替代 |
| 已在 Word 開啟／未儲存的文件 | 保留 exact live attachment；不能直接修改磁碟 OOXML |
| 最終目錄與 Word 分頁、PDF | 仍由 Word 原生引擎處理；HTML 預覽不能代替 |
| 科學數據重繪 | 仍用 Python matplotlib／numpy／pandas；OfficeCLI 負責嵌入圖檔 |

原 v8 有十處圖號 REF 快取落後於書籤，重建版已更新。原稿的 `EqJ_322` 則指向不存在的「含 λG、λSS 的總損失」公式；現存 Difficult-GradNorm 總損失式並不相同，因此沒有擅自把它接到另一個公式。交付版第 9 頁仍保留此原稿錯誤並在報告揭露。

研究資料方面，ANN 3,024／3,219 筆資料版本關係、單一 seed 42、KCI 的 ANN 參考性質、尚未進行的閉迴路驗證，均保留原稿的限制敘述。本次文件工具驗證不消除這些研究限制。

## 重跑入口

在本專案 PowerShell 7 中：

```powershell
. ./scripts/env.ps1
./scripts/prepare_chat_replay.ps1
./scripts/replay_all_stages.ps1
./scripts/verify_replay_stages.ps1
.venv/Scripts/python.exe -X utf8 scripts/audit_replay_numbers.py
.venv/Scripts/python.exe -X utf8 scripts/test_project_mcp.py
```

原生 Word 步驟需能存取桌面登入工作階段；受限環境的 `0x80070520` 並不代表文件已經通過 QA。科學繪圖的來源與適配程式保留在 `plots/scripts/`；完整版本清單在 `evidence/runtime-freeze.txt`。輸出檔存在時不覆寫，先查看既有紀錄再續跑。

此報告與 skills 保存本次操作經驗；依本次所有檔案留在專案的要求，沒有修改全域 memory 檔案。
