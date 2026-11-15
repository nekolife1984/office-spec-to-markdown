#!/usr/bin/env python3
"""Generate sample .docx / .xlsx specification documents for testing.

The samples deliberately include the hard parts: headings, a table and an
image in Word; merged cells, a table, a cell comment, an anchored image and
a text box (injected drawing XML) in Excel.

Usage: python3 make_samples.py <output_dir>
Requires: python-docx, openpyxl, Pillow
"""
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from PIL import Image, ImageDraw


def make_png(path):
    img = Image.new("RGB", (320, 200), "white")
    d = ImageDraw.Draw(img)
    d.rectangle([10, 10, 310, 190], outline="navy", width=3)
    d.rectangle([30, 60, 140, 120], outline="teal", width=2)
    d.rectangle([180, 60, 290, 120], outline="orange", width=2)
    d.line([140, 90, 180, 90], fill="black", width=2)
    img.save(path)
    return path


def make_docx(out_dir, png):
    from docx import Document
    doc = Document()
    doc.add_heading("サンプル仕様書", level=1)
    doc.add_paragraph("これは変換テスト用のWord仕様書です。図表を含みます。")
    doc.add_heading("画面概要", level=2)
    doc.add_paragraph("ログイン画面の仕様を以下に示します。")
    table = doc.add_table(rows=3, cols=3)
    data = [["項目", "必須", "備考"], ["ユーザーID", "○", "メール形式"], ["パスワード", "○", "8文字以上"]]
    for i, row in enumerate(data):
        for j, v in enumerate(row):
            table.rows[i].cells[j].text = v
    doc.add_paragraph("図1：画面遷移のイメージ")
    doc.add_picture(str(png))
    path = out_dir / "sample.docx"
    doc.save(path)
    return path


TEXTBOX_DRAWING = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
 <xdr:twoCellAnchor>
  <xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:sp>
   <xdr:nvSpPr><xdr:cNvPr id="2" name="テキストボックス 1"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>
   <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm>
   <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
   <xdr:txBody><a:bodyPr/><a:lstStyle/>
    <a:p><a:r><a:t>補足：パスワードは8文字以上必須です</a:t></a:r></a:p>
   </xdr:txBody>
  </xdr:sp>
  <xdr:clientData/>
 </xdr:twoCellAnchor>
</xdr:wsDr>
"""


def inject_textbox(xlsx_path):
    """Append a text-box anchor into the first sheet's existing drawing part.

    openpyxl already created xl/drawings/drawing1.xml for the sample image,
    so the text box is merged into that part instead of adding a new one.
    """
    xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    ET.register_namespace("xdr", xdr)
    ET.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
    anchor = ET.fromstring(TEXTBOX_DRAWING).find(f"{{{xdr}}}twoCellAnchor")
    tmp = xlsx_path.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(xlsx_path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.namelist():
            data = zin.read(item)
            if item == "xl/drawings/drawing1.xml":
                root = ET.fromstring(data)
                root.append(anchor)
                data = ET.tostring(root, xml_declaration=True, encoding="UTF-8")
            zout.writestr(item, data)
    tmp.replace(xlsx_path)


def make_xlsx(out_dir, png):
    from openpyxl import Workbook
    from openpyxl.comments import Comment
    from openpyxl.drawing.image import Image as XLImage

    wb = Workbook()
    ws = wb.active
    ws.title = "画面仕様"
    ws.merge_cells("B2:D2")
    ws["B2"] = "ログイン画面仕様"
    headers = ["No", "項目", "必須", "備考"]
    for j, h in enumerate(headers, start=2):
        ws.cell(row=4, column=j, value=h)
    rows = [[1, "ユーザーID", "○", "メール形式"], [2, "パスワード", "○", "8文字以上"]]
    for i, r in enumerate(rows, start=5):
        for j, v in enumerate(r, start=2):
            ws.cell(row=i, column=j, value=v)
    ws["C5"].comment = Comment("入力チェックはサーバー側でも実施", "spec")
    ws.add_image(XLImage(str(png)), "B8")
    ws2 = wb.create_sheet("コード表")
    ws2.append(["コード", "名称"])
    ws2.append(["01", "有効"])
    ws2.append(["02", "無効"])
    path = out_dir / "sample.xlsx"
    wb.save(path)
    inject_textbox(path)
    return path


def main():
    out_dir = Path(sys.argv[1] if len(sys.argv) > 1 else "samples")
    out_dir.mkdir(parents=True, exist_ok=True)
    png = make_png(out_dir / "_figure.png")
    print(make_docx(out_dir, png))
    print(make_xlsx(out_dir, png))


if __name__ == "__main__":
    main()
