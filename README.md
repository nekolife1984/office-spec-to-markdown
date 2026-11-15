# Officeスキル：Word/Excelを読んで、Markdownに変換する

Word（.docx）・Excel（.xlsx）をAIが読むためのスキルと、仕様書・設計書を
Markdown（GFM）へ変換するためのスキル、2本セットです。

「読む」ことと「変換する」ことを分けた、2つのエージェントスキルです：

| スキル | 役割 | 使う場面 |
|---|---|---|
| **office-docs-reader** | AIが.docx/.xlsxを読むために必要なものをすべて抽出する：本文、表、セル、図形・テキストボックスのテキスト、コメント、位置付きの画像、個数 | AIにWord/Excelファイルを読ませる・検索させる・要約させる・質問に答えさせるとき |
| **office-to-markdown** | readerの抽出物を、明示的なルールで保守しやすいMarkdownに再構成し、個数で検証する | Word/Excel文書（特に仕様書・設計書）をGit／docs-as-code向けの.mdファイルにするとき |

## なぜ「読む」と「変換する」を分けるのか

一発変換では、仕様書で一番大事な部分が黙って落ちます：図形や吹き出しの
中のテキスト、セルコメント、グラフ、セル結合の構造。「読む」だけの用途
（質疑応答、検索、要約）も変換なしで役に立ちます。そこで：

1. **office-docs-reader** が機械的に、忠実に抽出します——
   `scripts/extract.py` が `extracted.md`、`inventory.json`、`media/`、
   `work/shapes.json` を生成。再構成はせず、黙って落とすこともなし。
   抽出できないもの（グラフ、SmartArt）は個数を数えて警告に記録します。
2. **office-to-markdown** がその抽出物を受け取り、
   `references/conversion-rules.md` に従って再構成します（Excelのシート
   種別分類、図のdrawio.svg化かPNGかの判断、色・グレーアウトの意味の
   明示化など）。その後、`scripts/verify.py` で棚卸しと照合します。

図の変換方針：箱・矢印・テキストの図は編集可能な **drawio.svg**
（diagrams.net埋め込みSVG）に描き直すのが既定で、写真・スクリーン
ショット・高密度な図・データのグラフなど難しいものは **PNG** のままに
します。

## 構成

```
office-docs-reader/
  SKILL.md
  scripts/extract.py                .docx/.xlsx → extracted.md + inventory.json + media/
  references/extraction-notes.md    各内容がファイル形式の内部のどこにあるか
office-to-markdown/
  SKILL.md
  scripts/verify.py                 inventory.jsonとのリンク・個数検証
  scripts/make_drawio_svg.py        JSONの図仕様から編集可能な.drawio.svgを生成
  references/conversion-rules.md    Word/Excel変換ルールの詳細
  templates/convert-spec.prompt.md  GitHub Copilot用プロンプトファイル（.github/prompts/）
examples/
  make_samples.py                   テスト用サンプル.docx/.xlsxの生成
```

## クイックスタート

```bash
# 1. 読む／抽出する（docx: pandocがあれば使用、なければ標準ライブラリの予備抽出
#    xlsx: pip install openpyxl Pillow）
python3 office-docs-reader/scripts/extract.py spec.xlsx -o out/
# → out/extracted.md、out/inventory.json、out/media/、out/work/shapes.json

# 2. 変換する：office-to-markdown/references/conversion-rules.md に従って
#    out/extracted.md を out/spec.md に再構成する（エージェントの仕事）

# 3. 棚卸しと照合して検証する
python3 office-to-markdown/scripts/verify.py out/ spec.md
```

## エージェントでの使い方

- **Agent Skills対応エージェント**（VS CodeのGitHub Copilot、Claude
  Code、Codexなど）：2つのスキルフォルダをスキル用ディレクトリにコピー
  するか、ワークスペースから参照してください。片方だけでも使えます
  （変換しないならreaderだけ入れる）。
- **GitHub Copilot用プロンプトファイル**：
  `office-to-markdown/templates/convert-spec.prompt.md` をリポジトリの
  `.github/prompts/convert-spec.prompt.md` にコピーし、Copilot Chat
  （エージェントモード）で `/convert-spec` を実行してください。
  2スキルを順に駆動します。

## テスト

```bash
python3 examples/make_samples.py /tmp/spec-samples
python3 office-docs-reader/scripts/extract.py /tmp/spec-samples/sample.xlsx -o /tmp/out-xlsx
python3 office-docs-reader/scripts/extract.py /tmp/spec-samples/sample.docx -o /tmp/out-docx
# extracted.md を同じフォルダ内の最終.mdに再構成した後に：
python3 office-to-markdown/scripts/verify.py /tmp/out-xlsx
```

## ライセンス

MIT
