---
name: jtice-format
description: "Prepare, format, or audit manuscripts and submission files for the Journal of the Taiwan Institute of Chemical Engineers (JTICE), using its official author guide. Use for 台灣化工期刊/JTICE 投稿格式、摘要、字數、圖表、公式、文獻、Highlights、圖文摘要與投稿附件檢查. Does not apply to the Chinese 化工會刊 or automatically reformat a thesis."
---

# JTICE 期刊格式

將 JTICE 官方規定轉成可執行的排版與投稿檢查。規則快照核對日期：2026-10-21–22（Asia/Taipei）。本 skill 是格式規範，不是寫作風格、研究品質或錄用保證。

2026-10-23 已用本機瀏覽器重新讀取官方 Maximum length、Reference Style 與 Submission checklist；4500 words、200-word Original Paper 摘要、display-item 規定及前六位作者加 et al. 均仍存在。一般網頁擷取 403 時可改用可用瀏覽器讀取，不把單一工具失敗寫成官方頁面無法查證。已刊論文的差異與適用界線見 [sources-and-conflicts.md](references/sources-and-conflicts.md)。

## 使用方式

1. 確認目標是 **Journal of the Taiwan Institute of Chemical Engineers**，ISSN 1876-1070；辨識文章類型與階段（初投、修訂或接受後）。有稿件便先檢查；缺少的資料列為待確認，不為格式稽核要求研究寫作大綱審批。
2. 主文、摘要、圖表、公式及文獻：讀 [references/manuscript-format.md](references/manuscript-format.md)。投稿附件、聲明與送件前檢查：讀 [references/submission-package.md](references/submission-package.md)。查證來源、處理互相矛盾的規定或更新數值：讀 [references/sources-and-conflicts.md](references/sources-and-conflicts.md)。全稿格式檢查需讀三份。
3. 真正準備投稿或更新這份 skill 時，重查官方 Guide for Authors 與相關政策；優先查文章類型、摘要、篇幅、圖檔、AI 政策和 EM 必填項目。不能上網時，可依快照完成檢查，但標示「依 2026-10-21–22 快照，未即時查證」。不要將第三方轉載或已刊 PDF 當作新規定。
4. 逐項回報 **符合／不符／待確認／不適用**，每一不符項目附實際位置、觀察值、規定與修正方式。無實際圖檔不能證實解析度；沒有附件不能證實附件已提交；估算字數不能稱為正式核准字數。
5. 使用者要求修改時，保留來源另存修訂檔，除非已授權原檔修改。先處理可判定的格式項目，作者姓名、貢獻、利益衝突、經費、授權與資料公開情況缺資訊時只列待確認，不捏造或替作者聲明。完成後檢查來源檔與實際輸出 PDF。

## 先盤點、再成批處理

1. 先判定這次是完整展開、翻譯、格式修訂或正式投稿。採使用者最新要求；目前 Claus 稿先完整寫作，字數與 display-item 僅統計，不先設各節預算或壓縮 Results。翻譯保留已確認的內容、表格、圖說、原生公式與引用。格式檢查不要求重新審批研究大綱。
2. 現有 DOCX 先執行 EndNote MCP `preflight_jtice_manuscript`；有來源稿便提供 `source_document_path`。建立來源論述—文獻對照與原生物件盤點；刪除文獻須有內容上的理由。提早找出低解析圖、須重畫的示意圖、未填齊書目、樣式漏印欄位和詞表同名重複，避免交付前才重做整份。
3. 新文獻用 `import_verified_endnote_reference`：使用者提供名稱／DOI／含 DOI 連結／可讀 PDF，唯一核對、去重後原生匯入。標題太短或重名時補作者／年份／DOI，不猜第一個搜尋結果。不把匯入成功當成 JTICE 樣式合格；用 `diagnose_endnote_journal_abbreviations` 區分記錄、詞表與實際 CWYW 輸出。
4. 多筆引用先排好句子錨點，用 `batch_edit_endnote_citations` 在一次 Word 工作階段內完成並統一更新，不逐筆啟動 Word。成功的環境探測在狀態未變時沿用。線上查證保存來源與本次結果；不要因相同疑問重查同一篇。工具失敗先看錯誤和文件／匯入狀態，改變已確認的故障條件後再試，避免空等與重複匯入。
5. 一個已授權編輯批次完成後，統一執行原生目錄／交互參照更新、字級驗證與渲染檢查。內容或版面再變才重做受影響的最終檢查，不能以中間多次渲染代替完成檔驗證。同一工作階段已即時查證並保存的官方規定可沿用；正式投稿、規則更新或發現衝突時才再查對應項目。

以下是期刊規範；在完整展開階段保存為之後投稿精簡的檢查基準。

## 期刊規範

- 英文；12 pt；雙行距；單欄。Times 或 Times New Roman 是官方偏好。不要把已刊雙欄版式當成投稿模板。
- Original Paper：文字上限 4500 words，明列包含主文、文獻及圖說；另有圖表 display-item 額度。Short Communication：3000 words；Letter：1200 words，兩者包含 display-item 換算。詳見篇幅表，不得簡化為所有稿件都 7500 words。
- Original Paper 摘要不超過 200 words，Short Communication 不超過 150 words；現行網頁要求三段結構。採 Background、Methods、Significant Findings，並保留官方標題用語不一致的記錄。
- 摘要後最多 6 keywords；Highlights 必交，3–5 點、每點含空白最多 85 characters，獨立可編輯檔。
- Graphical abstract 必交，獨立檔，至少高 531 × 寬 1328 px，縮至 5 × 13 cm 仍清楚。不要把高寬反過來。
- 引文是 **Vancouver 數字方括號制**，依首次引用順序編號；不是僅因使用方括號就套 IEEE。超過 6 位作者列前 6 位再 et al.。
- 表格及公式必須可編輯；圖、表、scheme 及圖說嵌入正文適當位置，圖檔另備獨立來源檔。圖說不能烘焙成圖中的像素。
- 不宣稱「Your Paper Your Way」免除期刊明列的檢查項目。現行網頁與舊 .doc checklist 的摘要段數不同，採網頁規定並揭露差異。

## DOCX 與本機工作區

圖表格式稽核不能只檢查是否套用樣式名稱。檢查圖像實際尺寸／白邊／圖中文字、caption 的完整文字與分頁、原生標號與本文交互參照、表格實際字級及跨頁情形。圖在上、可編輯圖說在下；表題在上，避免表題和表頭分離。按實際插入尺寸渲染，確認上下標無缺字、caption 不裁斷，短表沒有單列孤立到下一頁。版面間距是工作選擇，不將未明定數值寫成官方要求。

數字制引用放在所支持的句子或作者姓名旁；同段提及不同研究時逐一定位，不一律堆在段尾。Condense 計畫書時保留來源—論述對照，交代刪除與補回的文獻。核對真正產生的卷號、頁碼／文章編號、期刊縮寫與作者截斷，不以「選到 JTICE EndNote 樣式」代替書目驗證。學術搜尋及原生記錄維護交由可用的 EndNote 技能／工具，不直接修改其資料庫。

EndNote 的會議／書章卷號可能存在於記錄中而未被樣式輸出；另存樣式副本補足對應模板，再經 CWYW 驗證，保留原版。模板中的 DOI 網址常值與欄位連結符須沿用 EndNote 原生語法；至少分別檢查有 DOI 與無 DOI 的實際書目，不能留下 DOI 欄位名稱、反引號、豎線或空網址。期刊詞表同名重複項可能使縮寫失效，即使各項的 Abbreviation 1 已相同；依 EndNote 技能的備份、授權及唯一名稱表流程處理，不能只補空欄。英文修訂時須連同表格及 captions 翻譯，保留所有原生欄位，另附分節 Word 計數與計數範圍；Results 留白時也報目前總數與後續篇幅餘額。

實際編輯 Word 時遵守目標工作區的 AGENTS.md，並讀取可用的 word-editing／docx-text-ops／docx-design-ops 等相應技能。這些工具操作規範不等於 JTICE 要求。

在 `D:\word_editing_env_officeCLI`：使用其 `.venv\Scripts\python.exe`；Word 自動化前設定本機 TEMP/TMP；新增英文用 Times New Roman；中文工作註解依工作區用標楷體；一般內文及表格維持 12 pt，設定普通與複雜文字字級。依當地規定更新、驗證**既有**原生目錄與交互參照，執行字級檢查和視覺 QA。JTICE 指南未要求稿件新建論文目錄、圖目錄、表目錄或公式目錄，不因本機驗證程序而自動新增。

保留 Word 工作母檔的原生 EndNote/CWYW 欄位。官方要求提交稿移除文獻管理器 field codes；僅在準備實際提交檔時，另製出版副本並核對移除前後引文、文獻、公式與圖表編號一致。不要對母檔全面 unlink fields。

## 交付內容

按任務規模提供：修訂檔（若有）、有來源的格式檢查清單、篇幅與 display-item 計算口徑、附件缺件清單、未解決的官方衝突。尚有無法驗證的項目時，區分「已完成檔案修改」與「可正式投稿」。不要自動修改碩士論文、投稿系統、作者資料或代送稿件。

區分內容發展階段與正式投稿檢查。Results 尚未撰寫時，既有草稿的剩餘字數只表示當前算術差額，不能限制結果章只能用該餘額。使用者指定先完整展開時，先保留所需論述、方法細節與圖表，待內容完整後再精簡並按官方口徑檢查；必要且無法精簡的超額稿由作者取得編輯適用指示，已刊長文不構成默認例外。

本工作區目前的 Claus 英文期刊稿：使用者於 2026-10-23 明確指定「先不管字數跟圖表限制，先寫多，真的需要再修短」。目前不設各節字數預算，不因字數或圖表數量／等效篇幅預先刪減內容；先完整撰寫，之後另整理精簡投稿版。字數統計僅供參考，不據此縮減 Results 的空間。官方規範仍保留供投稿階段使用，引用正確性、圖表可讀性及證據界線繼續維持。
