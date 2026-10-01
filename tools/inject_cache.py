# -*- coding: utf-8 -*-
"""openpyxl로 만든 원본(.xlsx)의 수식은 그대로 두고, LibreOffice로 재계산한 사본의 값을 캐시(<v>)로 넣는다.

LibreOffice가 다시 저장한 파일은 '01_Inputs'! 같은 숫자 시작 시트명의 따옴표를 빼서 Excel에서 복구 오류가 나므로,
배포 파일은 항상 openpyxl 원본 + 캐시 값으로 만든다.
사용: python inject_cache.py <openpyxl 원본.xlsx> <LibreOffice 재계산본.xlsx> <out.xlsx>
"""
import datetime as dt
import re
import sys
import zipfile

import openpyxl
from lxml import etree

SRC, CALC, OUT = sys.argv[1:4]
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
q = lambda t: f"{{{NS}}}{t}"
EPOCH = dt.datetime(1899, 12, 30)

vals = openpyxl.load_workbook(CALC, data_only=True)
zin = zipfile.ZipFile(SRC)
wbx = etree.fromstring(zin.read("xl/workbook.xml"))
rels = etree.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
rid2target = {r.get("Id"): r.get("Target") for r in rels}
part2sheet = {}
for sh in wbx.iter(q("sheet")):
    tgt = rid2target[sh.get(f"{{{RNS}}}id")].lstrip("/")
    part2sheet[tgt if tgt.startswith("xl/") else "xl/" + tgt] = sh.get("name")

calc = wbx.find(q("calcPr"))
if calc is None:
    calc = etree.SubElement(wbx, q("calcPr"))
calc.set("fullCalcOnLoad", "1")

n_set = 0
zout = zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    if item.filename == "xl/workbook.xml":
        data = etree.tostring(wbx, xml_declaration=True, encoding="UTF-8", standalone=True)
    elif item.filename in part2sheet:
        ws = vals[part2sheet[item.filename]]
        root = etree.fromstring(data)
        for c in root.iter(q("c")):
            f = c.find(q("f"))
            if f is None:
                continue
            v = ws[c.get("r")].value
            for old in c.findall(q("v")):
                c.remove(old)
            if v is None:
                continue
            if isinstance(v, bool):
                t, txt = "b", "1" if v else "0"
            elif isinstance(v, (int, float)):
                t, txt = None, repr(float(v)) if isinstance(v, float) else str(v)
            elif isinstance(v, dt.datetime):
                t, txt = None, repr((v - EPOCH).total_seconds() / 86400)
            elif isinstance(v, dt.date):
                t, txt = None, str((dt.datetime(v.year, v.month, v.day) - EPOCH).days)
            elif isinstance(v, str) and re.fullmatch(r"#(N/A|VALUE!|REF!|DIV/0!|NUM!|NAME\?|NULL!)", v):
                t, txt = "e", v
            else:
                t, txt = "str", str(v)
            if t:
                c.set("t", t)
            elif "t" in c.attrib:
                del c.attrib["t"]
            etree.SubElement(c, q("v")).text = txt
            n_set += 1
        data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    zout.writestr(item, data)
zout.close()
print("cached values:", n_set, "->", OUT)
