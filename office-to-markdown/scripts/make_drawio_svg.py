#!/usr/bin/env python3
"""Build a .drawio.svg file from a simple diagram description (JSON).

A .drawio.svg is an SVG image that also embeds the diagrams.net (draw.io)
mxGraph XML in the root element's `content` attribute, so it renders
anywhere SVG renders AND can be opened and edited in diagrams.net /
draw.io as a native diagram. This is the preferred output format for
figures converted by the office-to-markdown skill; use PNG instead when a
figure is too complex for this simple node/edge model.

Input JSON:
{
  "title": "optional, used as the diagram name",
  "nodes": [
    {"id": "a", "label": "開始", "shape": "ellipse"},
    {"id": "b", "label": "入力チェック", "shape": "rect", "x": 40, "y": 120},
    {"id": "c", "label": "OK?", "shape": "diamond"}
  ],
  "edges": [
    {"from": "a", "to": "b"},
    {"from": "b", "to": "c", "label": "完了"}
  ]
}

- shape: rect (default) | rounded | ellipse | diamond
- x/y are optional; nodes without positions are laid out in a grid
  (3 columns). Width/height default to 160x60 (diamond 120x80).
- Node ids must be unique; edge endpoints must exist.

Usage:
    python3 make_drawio_svg.py diagram.json -o figure.drawio.svg

The visible SVG rendering is generated from the same node/edge model as
the embedded mxGraph, so what readers see matches what draw.io opens.
"""
import argparse
import html
import json
import sys
from pathlib import Path
from xml.etree import ElementTree as ET

SHAPES = {
    "rect": ("", "rect"),
    "rounded": ("rounded=1;", "rect"),
    "ellipse": ("ellipse;", "ellipse"),
    "diamond": ("rhombus;", "diamond"),
}
DEFAULT_W, DEFAULT_H = 160, 60


def layout(nodes):
    """Assign x/y to nodes that lack them (grid, 3 columns)."""
    auto = [n for n in nodes if "x" not in n or "y" not in n]
    for i, n in enumerate(auto):
        col, row = i % 3, i // 3
        n["x"] = 40 + col * 220
        n["y"] = 40 + row * 140
    for n in nodes:
        n.setdefault("w", 120 if n.get("shape") == "diamond" else DEFAULT_W)
        n.setdefault("h", 80 if n.get("shape") == "diamond" else DEFAULT_H)
        n.setdefault("shape", "rect")
        if n["shape"] not in SHAPES:
            sys.exit(f"unknown shape: {n['shape']} (use rect|rounded|ellipse|diamond)")


def build_mxfile(spec, nodes, edges):
    mx = ET.Element("mxfile", {"host": "app.diagrams.net", "type": "device"})
    diagram = ET.SubElement(mx, "diagram", {"name": spec.get("title", "diagram")})
    model = ET.SubElement(diagram, "mxGraphModel", {
        "dx": "800", "dy": "600", "grid": "1", "gridSize": "10",
        "page": "1", "pageScale": "1"})
    root = ET.SubElement(model, "root")
    ET.SubElement(root, "mxCell", {"id": "0"})
    ET.SubElement(root, "mxCell", {"id": "1", "parent": "0"})
    for n in nodes:
        style, _ = SHAPES[n["shape"]]
        cell = ET.SubElement(root, "mxCell", {
            "id": str(n["id"]), "value": str(n["label"]), "vertex": "1", "parent": "1",
            "style": style + "whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#333333;"})
        ET.SubElement(cell, "mxGeometry", {
            "x": str(n["x"]), "y": str(n["y"]),
            "width": str(n["w"]), "height": str(n["h"]), "as": "geometry"})
    for i, e in enumerate(edges):
        cell = ET.SubElement(root, "mxCell", {
            "id": f"e{i}", "value": str(e.get("label", "")), "edge": "1", "parent": "1",
            "source": str(e["from"]), "target": str(e["to"]),
            "style": "edgeStyle=orthogonalEdgeStyle;rounded=0;html=1;endArrow=block;"})
        ET.SubElement(cell, "mxGeometry", {"relative": "1", "as": "geometry"})
    return ET.tostring(mx, encoding="unicode")


def build_svg(spec, nodes, edges, mx_xml):
    by_id = {str(n["id"]): n for n in nodes}
    width = max(n["x"] + n["w"] for n in nodes) + 40
    height = max(n["y"] + n["h"] for n in nodes) + 40
    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" '
        f'viewBox="0 0 {width} {height}" content="{html.escape(mx_xml, quote=True)}">',
        "<defs><marker id=\"arrow\" markerWidth=\"10\" markerHeight=\"10\" refX=\"8\" refY=\"3\" "
        "orient=\"auto\"><path d=\"M0,0 L8,3 L0,6 Z\" fill=\"#333333\"/></marker></defs>",
        '<rect x="0" y="0" width="100%" height="100%" fill="#ffffff"/>',
    ]
    for e in edges:
        a, b = by_id[str(e["from"])], by_id[str(e["to"])]
        x1, y1 = a["x"] + a["w"] / 2, a["y"] + a["h"] / 2
        x2, y2 = b["x"] + b["w"] / 2, b["y"] + b["h"] / 2
        parts.append(f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="#333333" marker-end="url(#arrow)"/>')
        if e.get("label"):
            parts.append(f'<text x="{(x1 + x2) / 2}" y="{(y1 + y2) / 2 - 4}" font-size="12" '
                         f'text-anchor="middle" fill="#333333">{html.escape(str(e["label"]))}</text>')
    for n in nodes:
        x, y, w, h = n["x"], n["y"], n["w"], n["h"]
        _, svg_shape = SHAPES[n["shape"]]
        if svg_shape == "ellipse":
            parts.append(f'<ellipse cx="{x + w / 2}" cy="{y + h / 2}" rx="{w / 2}" ry="{h / 2}" fill="#ffffff" stroke="#333333"/>')
        elif svg_shape == "diamond":
            parts.append(f'<polygon points="{x + w / 2},{y} {x + w},{y + h / 2} {x + w / 2},{y + h} {x},{y + h / 2}" fill="#ffffff" stroke="#333333"/>')
        else:
            rx = 10 if n["shape"] == "rounded" else 0
            parts.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{rx}" fill="#ffffff" stroke="#333333"/>')
        parts.append(f'<text x="{x + w / 2}" y="{y + h / 2 + 4}" font-size="13" text-anchor="middle" '
                     f'fill="#000000">{html.escape(str(n["label"]))}</text>')
    parts.append("</svg>")
    return "\n".join(parts)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("input", help="diagram description JSON")
    ap.add_argument("-o", "--out", required=True, help="output .drawio.svg path")
    args = ap.parse_args()
    spec = json.loads(Path(args.input).read_text(encoding="utf-8"))
    nodes = spec.get("nodes", [])
    edges = spec.get("edges", [])
    if not nodes:
        sys.exit("no nodes in input")
    ids = [str(n["id"]) for n in nodes]
    if len(ids) != len(set(ids)):
        sys.exit("duplicate node id")
    for e in edges:
        if str(e["from"]) not in ids or str(e["to"]) not in ids:
            sys.exit(f"edge references unknown node: {e}")
    layout(nodes)
    mx_xml = build_mxfile(spec, nodes, edges)
    svg = build_svg(spec, nodes, edges, mx_xml)
    out = Path(args.out)
    out.write_text(svg, encoding="utf-8")
    print(f"wrote {out} ({len(nodes)} nodes, {len(edges)} edges)")


if __name__ == "__main__":
    main()
