"""扫描已被 Git 跟踪的 HTML，使用 Python 标准库生成两列表格，无需安装依赖。"""
import os
import subprocess
import zipfile
from pathlib import Path
from urllib.parse import quote
from xml.sax.saxutils import escape

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "刷题链接目录.xlsx"
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def generate():
    repo = os.environ.get("GITHUB_REPOSITORY", "NNNewGiter/guide-exam-online")
    owner, name = repo.split("/", 1)
    base = os.environ.get("PAGES_BASE_URL") or (
        f"https://{owner.lower()}.github.io/" +
        ("" if name.lower() == f"{owner.lower()}.github.io" else quote(name) + "/")
    )
    raw = subprocess.check_output(["git", "-c", f"safe.directory={ROOT.as_posix()}",
                                   "ls-files", "-z"], cwd=ROOT)
    paths = sorted((p for p in raw.decode("utf-8").split("\0")
                    if p.lower().endswith(".html") and (ROOT / p).is_file()), key=str.casefold)
    rows = [(Path(p).stem, base.rstrip("/") + "/" + quote(p, safe="/")) for p in paths]
    def cell(ref, text, style):
        return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t>{escape(text)}</t></is></c>'
    body = '<row r="1" ht="26" customHeight="1">' + cell("A1", "标题", 1) + cell("B1", "链接", 1) + '</row>'
    links, relations = [], []
    for i, (title, url) in enumerate(rows, 2):
        body += f'<row r="{i}" ht="42" customHeight="1">' + cell(f"A{i}", title, 0) + cell(f"B{i}", url, 2) + '</row>'
        links.append(f'<hyperlink ref="B{i}" r:id="link{i}"/>')
        relations.append(f'<Relationship Id="link{i}" Type="{REL}/hyperlink" Target="{escape(url, {chr(34): "&quot;"})}" TargetMode="External"/>')
    parts = {
        "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
        "_rels/.rels": f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="book" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": f'<workbook xmlns="{NS}" xmlns:r="{REL}"><sheets><sheet name="刷题链接" sheetId="1" r:id="sheet"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="sheet" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="style" Type="{REL}/styles" Target="styles.xml"/></Relationships>',
        "xl/styles.xml": f'<styleSheet xmlns="{NS}"><fonts count="3"><font><sz val="11"/><name val="Arial"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Arial"/></font><font><color rgb="FF1765B5"/><u/><sz val="11"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF203D60"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="3"><xf fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf fontId="1" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf><xf fontId="2" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>',
        "xl/worksheets/sheet1.xml": f'<worksheet xmlns="{NS}" xmlns:r="{REL}"><dimension ref="A1:B{len(rows)+1}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols><col min="1" max="1" width="57" customWidth="1"/><col min="2" max="2" width="105" customWidth="1"/></cols><sheetData>{body}</sheetData><autoFilter ref="A1:B{len(rows)+1}"/><hyperlinks>{"".join(links)}</hyperlinks></worksheet>',
        "xl/worksheets/_rels/sheet1.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ''.join(relations) + '</Relationships>',
    }
    # 固定 ZIP 时间戳，内容未变时不会产生多余提交。
    with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as z:
        for path, xml in parts.items():
            info = zipfile.ZipInfo(path, (2020, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + xml)
    print(f"已生成 {OUTPUT.name}，共 {len(rows)} 个 HTML 链接（不含 README）。")


if __name__ == "__main__":
    generate()
