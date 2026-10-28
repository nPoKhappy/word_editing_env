# EndNote Word 技能操作說明

這個本機技能採用 [felimet/endnote-for-word](https://github.com/felimet/endnote-for-word) 作為基礎，透過 Word COM 呼叫 EndNote 的 Cite While You Write（CWYW），不使用 Computer Use。上游版本、原生匯入就緒判定的相容性修補與授權記錄在 `upstream.json`，MIT 授權保留於 `../vendor/endnote-for-word/LICENSE`。1.5.0 的線上查證／自動匯入、批次引用、縮寫診斷、PDF 搜尋及現成套件整合見 [extensions.md](extensions.md)。

## 安裝位置與啟動

- 技能：`D:\word_editing_env_officeCLI\.agents\skills\endnote-word`
- MCP 入口：`scripts\endnote_mcp.py`
- 本機擴充：`scripts\local_endnote.py`、`scripts\edit-citation.ps1`
- 上游程式與測試：`vendor\endnote-for-word`
- 工作區 MCP 設定：`D:\word_editing_env_officeCLI\.codex\config.toml`

此專案 MCP 設定使用本 checkout 的執行環境、腳本與狀態快照；本次沒有修改全域 MCP 設定。更改專案設定後須重新載入連線，既有對話中的服務不會自動切換。

原生 MCP 名稱為 `endnote-word-local`，另登錄 `endnote-library-search` 與 `endnote-reference-check`。原生橋接不需額外 MCP SDK；擴充的搜尋／查證套件獨立安裝於 `.tools/endnote`，不需外部帳號或 API key。線上查證會送出選定書目欄位至公開索引；本機全文／語意搜尋使用快取。需安裝桌面版 Word、桌面版 EndNote 及啟用 CWYW。原生欄位由 EndNote 產生，EndNote 必須能找到指定文獻庫。讀取群組只會操作暫存資料庫副本。

可直接執行單一工具。JSON 檔使用 UTF-8；JSON 的 Windows 路徑反斜線需寫成 `\\`，也可使用 `/`。

```powershell
$env:TEMP = 'D:\word_editing_env_officeCLI\.tmp'
$env:TMP = $env:TEMP
$env:ENDNOTE_WORD_POWERSHELL = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
$env:ENDNOTE_WORD_TIMEOUT_SECONDS = '120'
& 'D:\word_editing_env_officeCLI\.venv\Scripts\python.exe' -X utf8 `
  'D:\word_editing_env_officeCLI\.agents\skills\endnote-word\scripts\endnote_mcp.py' `
  --call search_endnote_group --arguments 'D:\word_editing_env_officeCLI\.tmp\search.json'
```

省略 `--call` 時，程式以 stdio MCP 模式運作。stdout 只傳輸 JSON；程序紀錄送到 stderr。CLI 的退出碼 0 表示成功，1 表示失敗；查看回傳的 `ok`、`error` 與驗證結果。

## 群組搜尋

`list_endnote_groups`：

```json
{"library_path":"D:/My_Research_References.enl"}
```

`search_endnote_group`：

```json
{
  "library_path":"D:/My_Research_References.enl",
  "group":"Claus",
  "title":"Advanced process control for sulfur recovery units",
  "limit":25
}
```

也可用 `group_id` 取代 `group`。兩者只能選一個；名稱採完整、大小寫不敏感匹配，不會找不到群組就退回全庫搜尋。其他篩選有 `query`、`author`、`year`、`doi`；`limit` 為 1–200。垃圾桶紀錄不會返回。

已確認這台電腦的 `Claus` 為手動群組，ID 9、8 筆會員，目標文獻的紀錄編號是 55。編號屬於該文獻庫，換庫或重新匯入後必須重新搜尋；不要把 55 寫死到通用程式。

群組儲存格式依實際 EndNote 2025 資料核對：`groups.spec` 為 XML，`TYPE;3` 表示手動群組；`members` 的版本標頭為 `00 00 00 02`，會員數與會員 ID 為 little-endian uint32。程式檢查版本、數量、長度與重複 ID；不認識的格式會報錯。智慧群組、組合群組、線上搜尋群組的規則不會自行推測或評估；EndNote XML 匯出也不適用群組操作。

## 插入正式引用

`insert_endnote_citation_after_paragraph`：

```json
{
  "input_path":"D:/word_editing_env_officeCLI/thesis_outline/改良型克勞斯製程論文章節與研究流程規劃_修正版_3_範圍定義更正.docx",
  "output_path":"D:/word_editing_env_officeCLI/thesis_outline/計畫書_新增引用.docx",
  "library_path":"D:/My_Research_References.enl",
  "group":"Claus",
  "rec_number":55,
  "expected_title":"Advanced process control for sulfur recovery units by artificial intelligence (AI) surrogate model",
  "paragraph_text":"引用代表性原始研究，連結尾氣品質估測。"
}
```

直接更新指定原檔時，在參數加入 `"overwrite": true` 並省略 `output_path`，或將 `output_path` 設成與 `input_path` 相同。預設仍須指定不存在的新輸出檔；`overwrite=true` 不允許覆寫另一個既有文件。移除引用工具也支援相同選項。

工具會在內部暫存副本執行 CWYW，核對引用身分、欄位數、其他正文與來源檔雜湊。驗證通過後，將原檔完整備份至同資料夾的 `.endnote-backups`，再以原檔名發布結果，不留下額外的編輯版 DOCX。回傳 `in_place`、`backup_path` 及 `source_unchanged`。驗證失敗不修改原檔；原檔被其他程序修改或鎖定時拒絕發布，不強制關閉使用者的 Word。更新 MCP 後須重新載入，既有對話的工具參數可能仍是舊版；可使用前述 CLI 單次呼叫新參數。

程式核對文獻庫中的編號與完整標題，以及可選的群組歸屬，再在唯一匹配段落的最後插入暫時引用。`ENFormatBibliography` 將其轉為正式 `ADDIN EN.CITE`，並產生 `ADDIN EN.REFLIST`。標題、作者、年份及紀錄編號必須符合真實文獻，才會發布結果。

上述舊工具的段落定位使用包含匹配，實際插入點是整段最後。句中引用請改用 `insert_endnote_citation_at_text`，不要將段尾工具當作句中插入。來源須為已有且沒有未處理追蹤修訂的 DOCX，目標須為正文、非表格段落。

## 句中插入及移動（1.2.0）

`insert_endnote_citation_at_text` 沿用 input/output、library、record、title 參數，另要求 `anchor_text`，並接受 `position: "after"`（預設）或 `"before"`。`paragraph_text` 必須唯一對應一段，`anchor_text` 必須在該段唯一且不與欄位重疊。例如將 `anchor_text` 設為 `Chou et al.`，可將該文獻置於作者姓名之後。數字引用前後的英文空白及標點由正文負責。

`move_endnote_citation_to_text` 不需要 library_path；以既有 record/title 識別單筆原生欄位並移到指定 anchor。它搬移完整的原生 CWYW 欄位，之後用 CWYW 重建文獻表。合併多筆欄位會拒絕；已在指定位置則驗證後不移動。`remove_preceding_space` 預設 true，只處理舊插入位置的一個分隔空白。

兩者沿用交易式驗證與 `overwrite=true` 自動備份。全段提到不同研究時應逐篇定位，不能把全部引文集中在段尾。MCP 重新載入後才顯示新增工具；單次 CLI 可立即呼叫。

## 書目欄位及期刊縮寫維護

本 MCP 現已提供線上名稱／DOI 查證和原生自動匯入新記錄；`import_verified_endnote_reference` 會核對身分、去重並回讀真實記錄編號。現有記錄的欄位及期刊詞表仍沒有已驗證的原生自動回寫 API。使用內建查證／診斷工具核對出版社、DOI、卷號、頁碼或文章編號，再經授權的 EndNote 原生介面修改；不得直接寫 SQLite，也不要匯入另一篇同名記錄取代核對。缺 DOI 不阻止引用。

期刊縮寫同時取決於 output style 的 Journal Names 設定及 Journals term list。匯入詞表後須檢查同名舊詞條是否仍有空白 Abbreviation 1；同名重複詞條即使各自填上相同縮寫，仍可能使 CWYW 保留全名。先備份完整四欄詞表及合併預覽，依適用的 UI 確認規則取得刪除／重建授權，再用 EndNote 原生匯入唯一名稱表。以唯讀快照核對名稱唯一、記錄數不變，最後驗證 Word 實際輸出；不能只因詞表匯入成功就宣稱縮寫修好。會議、書章及預印本依其真實類型核對，不為填滿欄位捏造卷號或頁碼。

既有文獻 DOI 未填仍可引用。程式不補寫 DOI、不新增重複文獻，只有明確設定 `overwrite=true` 才更新來源文件。保留文件目前的引用樣式；若要更換樣式，另用上游 `list_endnote_styles`、`set_endnote_citation_style`，依使用者要求操作。

## 移除指定引用

`remove_endnote_citation`：

```json
{
  "input_path":"D:/word_editing_env_officeCLI/thesis_outline/計畫書_新增引用.docx",
  "output_path":"D:/word_editing_env_officeCLI/thesis_outline/計畫書_移除引用.docx",
  "rec_number":55,
  "expected_title":"Advanced process control for sulfur recovery units by artificial intelligence (AI) surrogate model",
  "paragraph_text":"引用代表性原始研究，連結尾氣品質估測。",
  "remove_preceding_space":true
}
```

核對段落、編號、標題後，刪除完整的單筆引用欄位。仍有其他引用時，由 CWYW 重建编号與文末清單；最後一筆移除後，直接移除已無用途的文末清單欄位與 `EN.Libraries` 文件變數，避免零筆引用的格式化卡住。可能留下文末空段落，但其他正文與表格不會刪除。

`remove_preceding_space` 預設 `false`；只有確定引用前的那一個空白屬於插入分隔符時才設為 `true`。它不會整段去除空白或改寫原文。若一個引用欄位含多篇文獻（例如 `[1,2]`），工具會拒絕刪除，以免連帶移除其他文獻。這與 EndNote 的「Remove Field Codes」不同，後者會把整份文件轉為一般文字；本技能不呼叫該操作。

## 檢查與其他上游功能

`inspect_endnote_document` 接受 `document_path`，直接讀取 DOCX 儲存的欄位資料，返回引用數、各引用的紀錄／標題／作者／年份，以及文末清單文字。它不開啟 Word、不更新欄位、不驗證頁面版面。

保留的上游工具包括 `probe_endnote_word`、樣式列舉、文獻搜尋、文獻取得、manifest 預檢、marker 批次產生正式引用、欄位更新、樣式切換、取消格式化，以及 RIS 匯入／匯出。詳細參數由 MCP 的 `tools/list` 提供；上游操作說明在 `../vendor/endnote-for-word/README-zhTW.md`。

批次 marker／manifest 流程仍要求 DOI；本機既有紀錄引用流程沒有這個限制。不要呼叫 `configure_endnote_defaults` 或匯入 RIS，除非使用者要求修改樣式預設或文獻庫。

## 逾時與恢復

預設工作區設定讓 Word 操作最多執行 120 秒，MCP 工具上限 150 秒。較大型批次操作需要依規模調整這兩個時間；上游自動估算會被明確的 `ENDNOTE_WORD_TIMEOUT_SECONDS` 覆蓋。

每次操作建立隱藏、獨立的 Word 程序，並從其視窗取得實際 PID。逾時清理只終止橋接程式明確回報且不在原有程序集合中的 Word PID；不終止 EndNote 或其他 Word 程序。無法確認 PID 時，只回報失敗與程序資訊。未通過驗證或逾時不會發布輸出 DOCX。

若出現 `0x80070520`，先辨別是否為執行環境限制；已有授權時可使用適用的升權執行。`RPC 0x800706BE` 或 Word 崩潰則屬於 Word／CWYW 的實際失敗；不要反覆呼叫同一巨集或宣稱完成。來源文件保持原樣，先核對程序、引用樣式及文獻庫再重試。

## 驗證与維護

```powershell
& 'D:\word_editing_env_officeCLI\.venv\Scripts\python.exe' -X utf8 -m unittest discover `
  -s 'D:\word_editing_env_officeCLI\.agents\skills\endnote-word\tests' -v
& 'D:\word_editing_env_officeCLI\.venv\Scripts\python.exe' -X utf8 -m unittest discover `
  -s 'D:\word_editing_env_officeCLI\.agents\skills\endnote-word\vendor\endnote-for-word\tests' `
  -p 'test_*.py'
```

上游網路與原生 Office 整合測試在未明確啟用時會略過；離線測試通過不能取代實際 Word 操作。這次的實際驗證另記於 `validation.md`。更新上游時先保留版本／MIT 授權，檢查匯入模組的 API 相容性，再跑離線和文件副本測試。

若要停用，需同時移除使用者與工作區 `config.toml` 中的 `[mcp_servers.endnote-word-local]` 與其 env 子表，保留其他設定。只移除專案表格仍會載入使用者層級的登錄。技能資料夾可另外封存；不需更動 EndNote 文獻庫或 Office 外掛。
