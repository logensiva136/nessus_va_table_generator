from csv import DictReader
from docx.oxml.shared import OxmlElement, qn

def read_csv(path):
    with open(path, encoding="utf-8") as f:
        yield from DictReader(f)

# custom subnet extractor
def subnets(data: list[str]):
    subnet = set(host.split(".")[2] for host in data)
    return ", ".join([f"10.X.{i}.0/24" for i in subnet])


# set cell shade color
def shade_cells(cells, shade):
    """
Shade Colors:
\nDarkRed - Critical
\nRed - High
\nOrange - Medium
"""
    for cell in cells:
        tcPr = cell._tc.get_or_add_tcPr()
        tcVAlign = OxmlElement("w:shd")
        tcVAlign.set(qn("w:fill"), shade)
        tcPr.append(tcVAlign)

# subnets =["40","50","120","121"]

def read_csv(path):
    with open(path, encoding="utf-8") as f:
        yield from DictReader(f)

