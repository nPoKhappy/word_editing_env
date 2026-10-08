# 原生 Word 操作要點

以下為實際 Word COM 工作流程，不是用 python-docx 手造標號的替代方案。Word API 說明可參考 [Range.InsertCaption](https://learn.microsoft.com/en-us/office/vba/api/word.range.insertcaption) 與 [Range.InsertCrossReference](https://learn.microsoft.com/en-us/office/vba/api/word.range.insertcrossreference)。

## 呼叫 InsertCaption

- Label 是實際 caption 標籤；Style 是段落樣式，兩者分開檢查。
- 若該標籤不存在，才使用 `word.CaptionLabels.Add(label)` 建立。已有標籤保留其 `IncludeChapterNumber`、`ChapterStyleLevel`、`NumberStyle` 等設定。
- COM 可省略參數使用 `pythoncom.Missing` 或省略具名參數，不能傳 `None` 代替 TitleAutoText。Position 的 Above 為 0、Below 為 1。

```python
picture.Range.InsertCaption(
    figure_label, figure_title, pythoncom.Missing, 1, False
)
table.Range.InsertCaption(
    table_label, table_title, pythoncom.Missing, 0, False
)
```

`InlineShape` 的段落格式透過 `picture.Range.ParagraphFormat` 取得，沒有 `picture.ParagraphFormat`。

取得新 caption 的完整 `Paragraph.Range` 後套用 `doc.Styles('圖')`／`doc.Styles('表')`，不只選取 title 字串套用字元外觀。

## 公式右側標號

1. 在預定公式附近使用原生 InsertCaption，以原稿的公式標籤與括號規則產生編號。
2. 在本次插入範圍內，以實際 SEQ 欄位找到其段落；不能假定它一定是佔位文字的 Previous() 段落。
3. 取 caption 完整標號的 `Range.FormattedText`（不含段落記號），移入「定位字元 → OMath → 定位字元 → 編號」的同一段落。刪除原始 caption 段落，保留一個 SEQ 欄位。
4. OMath BuildUp 後，透過 OMath.Range 定位，不依賴 Find 原始線性字串；Word 可能將變數轉成數學字元。
5. 套用完整「式」段落樣式。版心寬度為 PageWidth 減去左右邊界及適用的縮排／欄寬；定位點分別使用該寬度的一半及全寬。不改寫全文件樣式，只在必要時調整此新段落的定位點。
6. 檢查原生 SEQ 已移動且沒有遺失，更新後仍延續原稿公式序列。

## 內文原生交互參照

優先使用 Word `Range.InsertCrossReference`（或 Selection 對應方法），指定實際 ReferenceType、ReferenceKind、ReferenceItem 與 `InsertAsHyperlink=True`。

- 先以 `doc.GetCrossReferenceItems(label)` 取得現有 caption 清單，依新項目完整標題及所在位置選出正確 ReferenceItem。
- ReferenceItem 是清單項目索引，不是顯示的 caption 編號；章節編號、重設序列及被刪項目會讓兩者不同。
- `wdOnlyLabelAndNumber`（3）可取得標籤及編號；公式的括號顯示依原稿標籤處理。不要讓引用包含整條公式或冗長圖說。
- 若使用原生書籤配合 REF 欄位，必須指向已由 InsertCaption 產生的編號；驗證連結與更新行為，不能預填固定數字作為結果。

## 更新目錄與頁碼

```python
# SEQ / REF 等更新與引用檢查先完成，再更新目錄。
doc.Repaginate()
for i in range(1, doc.TablesOfContents.Count + 1):
    doc.TablesOfContents(i).Update()
for i in range(1, doc.TablesOfFigures.Count + 1):
    doc.TablesOfFigures(i).Update()
doc.Repaginate()
```

- 對物件集合未收錄的 TOC 欄位也要辨識並更新。`TOC \c "圖"` 與 `TOC \c "表"` 是自動圖表目錄，不是一般文字。
- 缺少相應目錄時，以 `TablesOfFigures.Add` 建立，Caption 使用實際標籤；自訂樣式目錄則沿用已存在的設計。
- 完整更新後，根據重新分頁結果再更新頁碼，確認正文與目錄一致；最多三輪收斂檢查。
- 不用 `doc.Fields.Update()` 無差別更新含 EndNote 的全部欄位；逐類更新所需欄位及目錄物件。
- DOCX 儲存後重新開啟，核對目錄項目／REF 的快取結果與實際位置，然後匯出 PDF 驗證版面。

## 開啟中的文件

- 檔案鎖定或背景對話框造成 COM Open 等待時，不關閉使用者的 Word。改用獨立副本檔名或處理可安全存取的已開文件。
- 自動化採用獨立的 `DispatchEx('Word.Application')`；清理時只關閉本次建立的文件與 Word 實例。
- 不以強制終止所有 WINWORD.EXE 解決等待。原稿與使用者未儲存的編輯必須保留。
