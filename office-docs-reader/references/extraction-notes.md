# 抽出ノート ― .docx / .xlsx の内部のどこに何が入っているか

どちらの形式も、XMLパーツをまとめたZIPアーカイブである。ライブラリが
表に出してこない内容がある場合は、「その内容は存在しない」と結論づけず、
パーツを直接読むこと。

## Word（.docx）

| 内容 | 場所 |
|---|---|
| 本文・見出し・表 | `word/document.xml`（スタイルは `word/styles.xml`） |
| 画像 | `word/media/*`、`a:blip/@r:embed` → `word/_rels/document.xml.rels` で参照 |
| 脚注・文末脚注 | `word/footnotes.xml`、`word/endnotes.xml` |
| コメント | `word/comments.xml` |
| ヘッダー・フッター | `word/header*.xml`、`word/footer*.xml` |
| テキストボックス・図形 | `document.xml` 内の描画要素（`w:txbxContent`）— pandocは大部分を落とす |
| SmartArt | `word/diagrams/*` と `document.xml` 内のフォールバック描画 |

## Excel（.xlsx）

| 内容 | 場所 |
|---|---|
| セルの値・結合範囲 | `xl/worksheets/sheetN.xml`（＋ `xl/sharedStrings.xml`） |
| シートの順序・名前・表示状態 | `xl/workbook.xml` ＋ `xl/_rels/workbook.xml.rels` |
| シートごとの描画・コメントへのリンク | `xl/worksheets/_rels/sheetN.xml.rels` |
| 図形・テキストボックス・吹き出し・画像（アンカーとテキスト） | `xl/drawings/drawingN.xml` — 図形は `xdr:sp`（`grpSp` 内のネストもあり）、テキストは `a:t`、位置はアンカーの `xdr:from` |
| セルコメント | `xl/comments/commentN.xml`（フォルダ名は `xl/comments*` 側） |
| 画像ファイル | `xl/media/*` |
| グラフ | `xl/charts/chartN.xml` — オブジェクトであり画像ではない。描画して読む |
| セル内画像 | セルの `vm` 属性 → `xl/metadata.xml` → `xl/richData/*` → `xl/media/*`（openpyxlは読み込まない） |

## トラブルシューティング

- **画像が抽出できたがアンカーが分からない**：内容で照合する。浮動画像の
  アンカーはopenpyxlの `ws._images` で取れる。画像バイトのハッシュ
  （`img._data()`）を取り、`xl/media/*` のファイルと照合する。ファイル名
  だけでは位置は分からない。
- **シートが見当たらない**：棚卸しの `sheet_state` を確認する。非表示・
  完全非表示（veryHidden）のシートも、ユーザーが作業用だと言うまでは
  実在の内容として扱う。
- **数値がおかしい**：日付・時刻セルはXML内ではシリアル値であり、
  openpyxlなどのライブラリがセルスタイルから変換している。数式セルは
  最後にExcelで保存したときのキャッシュ値を示すため、別ツールが作った
  ファイルでは古いことがある。
- **日本語が文字化けする**：XMLはUTF-8。文字化けはたいてい自作パーサが
  共有文字列を誤ったエンコーディングで読んだ場合に起きる。即席の文字列
  処理よりopenpyxl/pandocを優先する。
- **巨大なシート**：自由形式の仕様書シートは、使用範囲が大きく、書式だけ
  の空セルで埋まっていることがある。末尾の全空行・全空列を切り落として
  （extract.pyは実施済み）からLLMに渡す。
