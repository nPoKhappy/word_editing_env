# DOCX テキスト構造リファレンス

[English](docx-text-structure.md)

## DOCX は ZIP パッケージ

`.docx` は ZIP パッケージです。XML を確認する前に展開します。

```bash
python3 docx-text-ops/scripts/unpack.py document.docx unpacked/
```

主なファイル:

```text
word/document.xml          本文、表、テキストボックス
word/header*.xml           ヘッダー
word/footer*.xml           フッター
word/footnotes.xml         脚注
word/endnotes.xml          文末脚注
word/comments.xml          コメント
word/styles.xml            スタイル。テキストのみの変更では触らない
word/_rels/*.rels          リレーションシップ。テキストのみの変更では触らない
```

## テキストの場所

多くの表示テキストは run 内の `w:t` に入っています。

```xml
<w:p>
  <w:r>
    <w:t>本文テキスト</w:t>
  </w:r>
</w:p>
```

表、ヘッダー、フッター、コメント、脚注、文末脚注、テキストボックスも同じ構造です。

## split run 問題

見た目上 1 つの語句でも、書式や編集履歴により複数の run に分かれることがあります。XML 内で 1 つの文字列として検索できるとは限りません。

## 空白

先頭または末尾に空白がある `w:t` では `xml:space="preserve"` を維持します。

## 変更履歴

`w:ins` や `w:del` 内の run は編集しません。Word で変更履歴を承認または却下してから作業します。
