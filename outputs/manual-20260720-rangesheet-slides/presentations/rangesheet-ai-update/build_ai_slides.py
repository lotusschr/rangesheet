import zipfile, pathlib, re, shutil, xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

src = pathlib.Path('outputs/manual-20260720-rangesheet-slides/presentations/rangesheet-ai-update/source.pptx')
out = pathlib.Path('outputs/manual-20260720-rangesheet-slides/presentations/rangesheet-ai-update/output/rangesheet_streamlit_architecture_ai_updated.pptx')
out.parent.mkdir(parents=True, exist_ok=True)

P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
ET.register_namespace('p', P_NS)
ET.register_namespace('a', A_NS)
ET.register_namespace('r', R_NS)
ET.register_namespace('', REL_NS)

def emu(x):
    return str(int(round(x)))

def tx_body(text, size=24, color='1F2937', bold=False, align='l'):
    paras = []
    for raw in str(text).split('\n'):
        if raw == '':
            raw = ' '
        b = '<a:b/>' if bold else ''
        paras.append(
            f'<a:p><a:pPr algn="{align}"/><a:r><a:rPr lang="en-US" sz="{int(size*100)}">{b}'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="Aptos"/>'
            f'<a:ea typeface="Aptos"/></a:rPr><a:t>{escape(raw)}</a:t></a:r></a:p>'
        )
    return '<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>' + ''.join(paras) + '</p:txBody>'

def shape_xml(sid, name, x, y, w, h, text='', size=20, color='1F2937', bold=False, fill=None, line=None, radius=False, align='l'):
    prst = 'roundRect' if radius else 'rect'
    fill_xml = '<a:noFill/>' if fill is None else f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
    line_xml = '<a:ln><a:noFill/></a:ln>' if line is None else f'<a:ln w="12700"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>'
    return f'''
<p:sp>
  <p:nvSpPr><p:cNvPr id="{sid}" name="{escape(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>{fill_xml}{line_xml}</p:spPr>
  {tx_body(text, size=size, color=color, bold=bold, align=align)}
</p:sp>'''

def slide_xml(kicker, title, subtitle, cards, note=None):
    W, H = 12192000, 6858000
    shapes = []
    shapes.append(shape_xml(2, 'bg', 0, 0, W, H, fill='F7FAF9', line=None))
    shapes.append(shape_xml(3, 'accent bar', 0, 0, 260000, H, fill='2BBFA4', line=None))
    shapes.append(shape_xml(4, 'kicker', 620000, 410000, 2200000, 300000, kicker, size=13, color='2BBFA4', bold=True))
    shapes.append(shape_xml(5, 'title', 620000, 780000, 10400000, 720000, title, size=30, color='111827', bold=True))
    shapes.append(shape_xml(6, 'subtitle', 640000, 1550000, 10000000, 520000, subtitle, size=16, color='4B5563'))
    x0, y0 = 720000, 2350000
    card_w, card_h = 3300000, 1180000
    gap_x, gap_y = 300000, 330000
    sid = 10
    for i, (head, body, accent) in enumerate(cards):
        col = i % 3
        row = i // 3
        x = x0 + col * (card_w + gap_x)
        y = y0 + row * (card_h + gap_y)
        shapes.append(shape_xml(sid, f'card {i+1}', x, y, card_w, card_h, fill='FFFFFF', line='DDE7E3', radius=True)); sid += 1
        shapes.append(shape_xml(sid, f'card accent {i+1}', x, y, 90000, card_h, fill=accent, line=None)); sid += 1
        shapes.append(shape_xml(sid, f'card head {i+1}', x+220000, y+150000, card_w-420000, 260000, head, size=15, color='111827', bold=True)); sid += 1
        shapes.append(shape_xml(sid, f'card body {i+1}', x+220000, y+470000, card_w-420000, 540000, body, size=12, color='374151')); sid += 1
    if note:
        shapes.append(shape_xml(80, 'bottom note', 720000, 6100000, 10400000, 420000, note, size=12, color='6B7280', fill='E8F8F5', line=None, radius=True))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="{A_NS}" xmlns:r="{R_NS}" xmlns:p="{P_NS}">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    {''.join(shapes)}
  </p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>'''

slides = [
    ('AI IN DEVELOPMENT', 'AI helped make a complex real-data web app feasible within the timeline.',
     'The project moved beyond classroom examples: large CSV files, connected filters, editable grids, business rules, reports, roles, audit logs, and performance constraints had to work together.',
     [
        ('Time constraint', 'Manual coding alone would not have been enough to build and debug all requested features in the short development window.', '2BBFA4'),
        ('Real business data', 'The app handled large operational datasets where small logic errors could affect tables, reports, and user decisions.', 'F6D975'),
        ('Higher complexity', 'Changes in one cell needed to update status, cluster summaries, range architecture, autosave, chatbot behavior, and report output.', '3B82F6'),
        ('Unexpected issues', 'Performance bottlenecks, column-name mismatches, stale state, and UI reruns appeared during actual use.', 'EF4444'),
        ('AI as support', 'AI was used to explain errors, suggest fixes, compare approaches, and speed up iteration.', '8B5CF6'),
        ('Human judgment', 'Requirements, business rules, validation, and final decisions still had to be checked by the developer.', '10B981'),
     ], 'Key message: AI accelerated learning and delivery, but the logic was still guided by business requirements.'),
    ('AI WORKFLOW', 'AI supported coding, debugging, and performance improvement across the application.',
     'Instead of using AI as a one-click generator, it was used as a development partner during build-test-fix cycles.',
     [
        ('Code writing', 'Generated implementation drafts for Streamlit UI, grid behavior, chatbot logic, report sections, and export workflows.', '2BBFA4'),
        ('Debugging', 'Helped trace errors such as KeyError, undefined variables, stale session values, and dropdown state not persisting.', 'EF4444'),
        ('Business logic mapping', 'Translated user actions such as Delete Some, Delete All, New Some, and NewNew into consistent table and report calculations.', 'F6D975'),
        ('Performance thinking', 'Suggested caching, lighter recalculation, reduced reruns, and precomputed summaries to improve responsiveness.', '3B82F6'),
        ('UX iteration', 'Helped refine chatbot, login, report preview, export buttons, pinned tables, and full-screen views based on feedback.', '8B5CF6'),
        ('Learning loop', 'Each AI suggestion became a chance to understand why the code behaved that way, not just copy the answer.', '10B981'),
     ], None),
    ('LEARNING OUTCOME', 'The biggest learning was connecting code behavior with business meaning.',
     'The project required both technical debugging and an understanding of how range decisions affect item movement, planograms, and management reporting.',
     [
        ('System thinking', 'Learned to trace data from source files through filters, grids, status logic, cluster tables, and reports.', '2BBFA4'),
        ('Data quality awareness', 'Learned why missing values, different column names, and inconsistent formats can break downstream logic.', 'F6D975'),
        ('State management', 'Learned how autosave, refresh, role access, and edited dropdown states must remain consistent across pages and DGs.', '3B82F6'),
        ('Risk logic', 'Learned how business rules like Top 10% best seller and Tail 10% lowest seller can guide warnings and review points.', 'EF4444'),
        ('Communication', 'Learned to convert detailed app activity into manager-friendly reports and audit summaries.', '8B5CF6'),
        ('Practical coding', 'Learned by debugging real problems that were not fully predictable at the design stage.', '10B981'),
     ], None),
    ('VALUE CREATED', 'The final app turns range editing into a connected review workflow, not just a spreadsheet screen.',
     'Users can edit range decisions, see impact, ask the chatbot, export outputs, and send management reports while keeping auditability.',
     [
        ('Connected tables', 'Main table changes update status, cluster counts, range architecture, and report outputs.', '2BBFA4'),
        ('Decision support', 'Warnings highlight risky actions such as deleting best sellers or adding lowest sellers.', 'EF4444'),
        ('Faster review', 'Chatbot answers common range questions and supports quick action commands.', '8B5CF6'),
        ('Manager summary', 'Reports summarize DG changes, SKU impact, planogram impact, product movement, and sales risk.', 'F6D975'),
        ('Governance', 'Login roles and audit logs support controlled access and traceability.', '3B82F6'),
        ('Export readiness', 'Range output and report files can be generated for business handoff.', '10B981'),
     ], 'Outcome: AI helped bridge the gap between limited time, unfamiliar technical areas, and real business complexity.'),
]

with zipfile.ZipFile(src, 'r') as zin:
    names = zin.namelist()
    slide_nums = sorted(int(re.search(r'slide(\d+)\.xml$', n).group(1)) for n in names if re.match(r'ppt/slides/slide\d+\.xml$', n))
    max_slide = max(slide_nums)
    pres_xml = zin.read('ppt/presentation.xml')
    pres_rels_xml = zin.read('ppt/_rels/presentation.xml.rels')
    ct_xml = zin.read('[Content_Types].xml')
    blank_slide_rels = zin.read('ppt/slides/_rels/slide11.xml.rels')

    pres_root = ET.fromstring(pres_xml)
    rels_root = ET.fromstring(pres_rels_xml)
    ct_root = ET.fromstring(ct_xml)
    sldIdLst = pres_root.find(f'{{{P_NS}}}sldIdLst')
    max_id = max(int(e.attrib['id']) for e in sldIdLst.findall(f'{{{P_NS}}}sldId'))
    existing_rids = []
    for rel in rels_root.findall(f'{{{REL_NS}}}Relationship'):
        rid = rel.attrib.get('Id','')
        m = re.match(r'rId(\d+)$', rid)
        if m:
            existing_rids.append(int(m.group(1)))
    max_rid = max(existing_rids)

    skip = set(['ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', '[Content_Types].xml'])
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename in skip:
                continue
            zout.writestr(item, zin.read(item.filename))

        for i, slide in enumerate(slides, start=1):
            sn = max_slide + i
            rid = f'rId{max_rid + i}'
            sid = str(max_id + i)
            ET.SubElement(sldIdLst, f'{{{P_NS}}}sldId', {'id': sid, f'{{{R_NS}}}id': rid})
            ET.SubElement(rels_root, f'{{{REL_NS}}}Relationship', {
                'Id': rid,
                'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
                'Target': f'slides/slide{sn}.xml'
            })
            ET.SubElement(ct_root, f'{{{CT_NS}}}Override', {
                'PartName': f'/ppt/slides/slide{sn}.xml',
                'ContentType': 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
            })
            zout.writestr(f'ppt/slides/slide{sn}.xml', slide_xml(*slide).encode('utf-8'))
            zout.writestr(f'ppt/slides/_rels/slide{sn}.xml.rels', blank_slide_rels)

        zout.writestr('ppt/presentation.xml', ET.tostring(pres_root, encoding='utf-8', xml_declaration=True))
        zout.writestr('ppt/_rels/presentation.xml.rels', ET.tostring(rels_root, encoding='utf-8', xml_declaration=True))
        zout.writestr('[Content_Types].xml', ET.tostring(ct_root, encoding='utf-8', xml_declaration=True))

print(out.resolve())
