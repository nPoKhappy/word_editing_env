# DOCX デザイン編集レシピ

[English](design-recipes.md)

既存 `.docx` ファイルの見た目を調整するための代表的な手順です。ファイルは `docx-text-ops/scripts/unpack.py` で展開済みである前提です。

## レシピ 1: 本文フォントとサイズを変える

対象: `word/styles.xml`

1. `inspect_design.py` で現在の状態を確認する。
2. `Normal` スタイルと、必要に応じて `w:docDefaults` を更新する。
3. `w:sz` と `w:szCs` は同じ half-point 値にする。
4. 直接書式よりスタイル変更を優先する。

## レシピ 2: 見出しスタイルを変える

対象: `word/styles.xml`

見出しスタイル ID を探し、`w:rPr` と `w:pPr` だけを編集します。`w:basedOn`、`w:next`、アウトラインレベルは保持します。

## レシピ 3: テーマカラーを更新する

対象: `word/theme/theme1.xml`

`w:themeColor` を参照している要素はテーマカラー変更の影響を受けます。直接指定された `w:color w:val` は影響を受けません。

## レシピ 4: テーマフォントを更新する

対象: `word/theme/theme1.xml`

`majorHAnsi` や `minorHAnsi` などテーマ参照を使うスタイルに効きます。直接フォント名が入っているスタイルは `styles.xml` 側で変更します。

## レシピ 5: インライン上書きを扱う

対象: 必要な場合のみ `word/document.xml`

インライン書式はスタイルより優先されます。スタイルで統一したい場合だけ、明示的に direct formatting を除去します。段落スタイル、番号付け、セクションプロパティなど構造要素は削除しません。

## 検証

- `inspect_design.py` で意図した変更が見える。
- テキスト抽出結果が変わっていない。
- Word または LibreOffice で開ける。
- `.rels` ファイルを変更していない。
