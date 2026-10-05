# DOCX Design Operations Skill

[English](SKILL.md)

既存の `.docx` / `.docm` ファイルの視覚デザインを調整するスキルです。フォント、配色、見出しスタイル、段落間隔、テーマを変更しながら、テキスト、画像、表、リレーションシップを保持します。

## 使う場面

- 本文フォントやサイズを変更したい
- 見出しの色・サイズ・余白を整えたい
- テーマカラーやテーマフォントを変更したい
- 文書の見た目を統一したい

## 対象外

- テキスト内容の編集（`docx-text-ops` を使用）
- 新規文書の作成
- 変更履歴が有効な文書
- 画像や図形の編集

## 基本ワークフロー

1. 元ファイルをコピーする。
2. `python3 docx-design-ops/scripts/inspect_design.py working_copy.docx` で現在のデザイン状態を確認する。
3. `python3 docx-text-ops/scripts/unpack.py working_copy.docx unpacked/` で展開する。
4. `word/styles.xml` または `word/theme/theme1.xml` を編集する。バッチ変更には `apply_design.py` を使う。
5. 初めての文書では `python3 docx-design-ops/scripts/apply_design.py unpacked/ spec.json --dry-run` を先に実行する。
6. `python3 docx-text-ops/scripts/pack.py unpacked/ edited.docx` で再梱包する。
7. `inspect_design.py` と Word / LibreOffice の目視確認で結果を検証する。

## 安全ルール

- `w:basedOn` と `w:next` は変更しない。
- `.rels` ファイルは編集しない。
- デザイン指定は原則 `styles.xml` と `theme1.xml` に集約する。
- `word/document.xml` はインライン書式の除去が明示的に必要な場合だけ触る。
- `w:sz` を変更する場合は `w:szCs` も同じ値にする。

詳細は [docx-design-structure.ja.md](references/docx-design-structure.ja.md) と [design-recipes.ja.md](references/design-recipes.ja.md) を参照してください。
