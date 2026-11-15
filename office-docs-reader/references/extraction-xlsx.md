# 抽出ノート ― Excel（.xlsx）の内部構造

| 内容 | 場所 |
|---|---|
| セルの値・結合範囲 | `xl/worksheets/sheetN.xml`（＋ `xl/sharedStrings.xml`） |
| 罫線・塗り・フォント等の書式 | `xl/styles.xml`（各セルのスタイルインデックス経由）。罫線で図を描いている仕様書では意味を持つことがあるが、extract.py の抽出（extracted.md）には含まれない。必要ならPDF描画で確認する |
| シートの順序・名前・表示状態 | `xl/workbook.xml` ＋ `xl/_rels/workbook.xml.rels` |
| シートごとの描画・コメントへのリンク | `xl/worksheets/_rels/sheetN.xml.rels` |
| 図形・テキストボックス・吹き出し・画像（アンカーとテキスト） | `xl/drawings/drawingN.xml` — 図形は `xdr:sp`（`grpSp` 内のネストもあり）、テキストは `a:t`、位置はアンカーの `xdr:from` |
| セルコメント | `xl/comments/commentN.xml`（フォルダ名は `xl/comments*` 側） |
| 画像ファイル | `xl/media/*` |
| グラフ | `xl/charts/chartN.xml` — オブジェクトであり画像ではない。描画して読む |
| セル内画像 | セルの `vm` 属性 → `xl/metadata.xml` → `xl/richData/*` → `xl/media/*`（openpyxlは読み込まない） |

