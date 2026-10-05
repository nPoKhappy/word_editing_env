# DOCX Text Operations Skill

[English](SKILL.md)

既存の `.docx` / `.docm` / `.dotx` / `.dotm` ファイルのテキストだけを編集するスキルです。スタイル、画像、表、番号付け、リレーションシップは保持します。

## 使う場面

- 本文、表、ヘッダー、フッター、脚注、文末脚注、コメントのテキスト編集
- 翻訳、校正、用語統一、プレースホルダー削除
- 既存レイアウトを維持したまま内容だけ変更したい場合

## 対象外

- 新規文書の作成
- スタイル、テーマ、ページレイアウトの再設計（`docx-design-ops` を使用）
- 変更履歴が有効な文書
- 埋め込みオブジェクトの内部編集

## 基本ワークフロー

1. 元ファイルをコピーする。
2. `python3 docx-text-ops/scripts/extract.py working_copy.docx` で全文を確認する。
3. 変更履歴の警告が出たら、Word で承認または却下してから続行する。
4. `python3 docx-text-ops/scripts/unpack.py working_copy.docx unpacked/` で展開する。
5. `word/document.xml` や `word/header*.xml` など、対象の XML だけを最小限編集する。
6. `python3 docx-text-ops/scripts/pack.py unpacked/ edited.docx` で再梱包する。
7. 再度 `extract.py` で古い文字列が消え、新しい文字列が正しいことを確認する。

## 安全ルール

- `.rels` ファイルは編集しない。
- XML を正規表現だけで置換しない。
- `<w:ins>` / `<w:del>` など変更履歴の run は編集しない。
- 先頭・末尾に空白がある `<w:t>` では `xml:space="preserve"` を維持または追加する。

詳細は [docx-text-structure.ja.md](references/docx-text-structure.ja.md) と [text-edit-recipes.ja.md](references/text-edit-recipes.ja.md) を参照してください。
