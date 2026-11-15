# 抽出ノート ― Word（.docx）の内部構造

| 内容 | 場所 |
|---|---|
| 本文・見出し・表 | `word/document.xml`（スタイルは `word/styles.xml`） |
| 画像 | `word/media/*`、`a:blip/@r:embed` → `word/_rels/document.xml.rels` で参照 |
| 脚注・文末脚注 | `word/footnotes.xml`、`word/endnotes.xml` |
| コメント | `word/comments.xml` |
| ヘッダー・フッター | `word/header*.xml`、`word/footer*.xml` |
| テキストボックス・図形 | `document.xml` 内の描画要素（`w:txbxContent`）— pandocは大部分を落とす |
| SmartArt | `word/diagrams/*` と `document.xml` 内のフォールバック描画 |

