# DOCX デザイン構造リファレンス

[English](docx-design-structure.md)

## デザイン関連ファイル

`.docx` は ZIP パッケージです。デザイン変更の主な対象は以下です。

```text
word/styles.xml            段落・文字・表・リストのスタイル
word/theme/theme1.xml      テーマフォントと配色
word/document.xml          本文。インライン書式除去が必要な場合だけ触る
```

デザインだけの変更では `.rels` ファイルを編集しません。

## 2 層モデル

Word の見た目は、スタイル・テーマ定義と、本文中の直接書式の 2 層で決まります。直接書式がある場合はスタイルより優先されます。

## `word/styles.xml`

- `w:docDefaults`: 文書全体の既定値
- `w:style`: スタイル定義
- `w:rPr`: フォント、サイズ、色、太字、斜体など
- `w:pPr`: 段落間隔、行間、揃え、インデントなど

`w:basedOn` と `w:next` はスタイル継承と次段落の挙動に関わるため変更しません。

## `word/theme/theme1.xml`

- `a:fontScheme`: major/minor フォント
- `a:clrScheme`: `accent1` などのテーマカラー

テーマ参照のあるスタイルはテーマ変更で自動的に変わります。直接指定の色やフォントは変わりません。

## 単位

| プロパティ | 単位 | 例 |
|------------|------|----|
| `w:sz`, `w:szCs` | half-point | `24` = 12pt |
| `w:spacing w:before`, `w:after` | point の 1/20 | `240` = 12pt |
| `w:spacing w:line` + `auto` | point の 1/20 | `276` = 1.15 行 |
| `w:ind w:left`, `w:right` | point の 1/20 | `720` = 0.5 inch |
