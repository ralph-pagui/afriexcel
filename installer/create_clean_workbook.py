import os
import zipfile

def create_workbook(output_path):
    # Contenus XML sans aucun BOM (encodage UTF-8 pur)
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/>
  <Override PartName="/xl/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/>
</Types>'''

    dot_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

    wb_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="AfriExcel Copilot" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>'''

    wb_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="webextensions/taskpanes.xml"/>
</Relationships>'''

    sheet1_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>AFRIEXCEL COPILOT - BUSINESS PLAN BANCAIRE SYSCOHADA (20 FEUILLES)</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>L'Assistant IA s'ouvre automatiquement dans le volet a droite.</t></is></c></row>
  </sheetData>
</worksheet>'''

    taskpanes_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:taskpane dockstate="right" visibility="1" width="380" row="1" locked="0">
    <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
  </wetp:taskpane>
</wetp:taskpanes>'''

    tp_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="webextension1.xml"/>
</Relationships>'''

    addin_guid = "b29d4c78-5e4a-4b91-912a-8457fce2a912"
    webextension_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{{{addin_guid}}}">
  <we:reference id="{addin_guid}" version="1.0.0.0" store="developer" storeType="developer"/>
  <we:alternateReferences/>
  <we:properties>
    <we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>'''

    files = {
        "[Content_Types].xml": content_types.encode('utf-8'),
        "_rels/.rels": dot_rels.encode('utf-8'),
        "xl/workbook.xml": wb_xml.encode('utf-8'),
        "xl/_rels/workbook.xml.rels": wb_rels.encode('utf-8'),
        "xl/worksheets/sheet1.xml": sheet1_xml.encode('utf-8'),
        "xl/webextensions/taskpanes.xml": taskpanes_xml.encode('utf-8'),
        "xl/webextensions/_rels/taskpanes.xml.rels": tp_rels.encode('utf-8'),
        "xl/webextensions/webextension1.xml": webextension_xml.encode('utf-8'),
    }

    if os.path.exists(output_path):
        try:
            os.remove(output_path)
        except:
            pass

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for arcname, data in files.items():
            z.writestr(arcname, data)

if __name__ == '__main__':
    desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
    out1 = os.path.join(desktop, 'Assistant_IA_Excel.xlsx')
    out2 = os.path.join(desktop, 'AssistantIA_AfriExcel.xlsx')
    out3 = r'F:\cop\Assistant_IA_Excel.xlsx'
    for path in [out1, out2, out3]:
        create_workbook(path)
        print("Workbook generated with 0 BOM:", path)
