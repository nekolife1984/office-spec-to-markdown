---
mode: agent
description: ワークスペース内のWord/Excel仕様書を、office-docs-readerとoffice-to-markdownの2スキルでMarkdownに変換する
---

指定するOffice仕様書（このワークスペース内の.docxまたは.xlsxファイル）を、
2つのスキルを順に使ってMarkdownに変換してください：

1. まずoffice-docs-readerスキルで読み取り、再構成の前に棚卸しの個数を
   見せてください：
   `python3 <reader-skill-path>/scripts/extract.py <file> -o <output-dir>`
2. office-to-markdownスキルで抽出物を再構成してください。その際、
   `references/conversion-rules.md` に従うこと。特に：
   - 図形・テキストボックス・コメントのテキストは、それが属する
     セクションへ吸収する
   - Excelシートは変換前に種別を分類する
   - 図は既定で編集可能な.drawio.svg（diagrams.net）に描き直す。
     難しい図（写真、スクリーンショット、高密度な図、データのグラフ）は
     PNG参照のままにし、テキスト説明を付ける
   - セル結合のある表は、基本は結合値を全セルに伝播してパイプ表に
     する。見た目が重要な表だけHTML表で結合を保持する
   - 罫線で図を表現している箇所は表にせず、図としてdrawio.svgに
     描き直すか、難しいものはPNGにする
   - 表の行、図形テキスト、コメントを黙って落とさない
3. `python3 <converter-skill-path>/scripts/verify.py <output-dir>` を実行し、
   inventory.jsonと個数を照合してください。
4. 報告してください：何を変換したか、何が画像/PDFのままでその理由は
   何か、何を落としたか（レイアウト、色、ヘッダー・フッター）、
   判断した箇所はないか。

.docx/.xlsxのバイナリを直接、正解として読んではいけません。また、
私が確認するまで結果をコミットしないでください。
