import pathlib
import zipfile
import xml.etree.ElementTree as ET

source = pathlib.Path("outputs/manual-20260720-rangesheet-odp-edit/source.odp")
ns = {
    "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
}

with zipfile.ZipFile(source) as zf:
    root = ET.fromstring(zf.read("content.xml"))

pages = root.findall(".//draw:page", ns)
print("pages", len(pages))
for idx, page in enumerate(pages, 1):
    name = page.attrib.get(f"{{{ns['draw']}}}name", "")
    texts = []
    for node in page.findall(".//text:p", ns) + page.findall(".//text:span", ns):
        value = "".join(node.itertext()).strip()
        if value:
            texts.append(value)
    print(f"--- {idx} {name} ---")
    print(" | ".join(texts[:20])[:900])
