// Script pour créer un classeur Excel avec WebExtension embarquée
// Utilise la manipulation ZIP directe (xlsx = zip)
const fs = require('fs');
const path = require('path');

// Minimal XLSX with WebExtension using raw ZIP creation
// We need: [Content_Types].xml, _rels/.rels, xl/workbook.xml, xl/_rels/workbook.xml.rels,
//          xl/worksheets/sheet1.xml, xl/webextensions/webextension1.xml,
//          xl/webextensions/_rels/webextension1.xml.rels, xl/webextensions/taskpanes.xml

const ADDIN_ID = 'b29d4c78-5e4a-4b91-912a-8457fce2a912';
const ADDIN_URL = 'https://afriexcel.netlify.app/index.html';
const MANIFEST_URL = 'https://afriexcel.netlify.app/manifest.xml';

// Simple ZIP file creator (no compression, store only)
class SimpleZip {
  constructor() {
    this.files = [];
    this.centralDir = [];
    this.offset = 0;
  }

  addFile(name, content) {
    const nameBuffer = Buffer.from(name, 'utf8');
    const contentBuffer = Buffer.from(content, 'utf8');
    
    // CRC32
    const crc = this.crc32(contentBuffer);
    
    // Local file header
    const header = Buffer.alloc(30 + nameBuffer.length);
    header.writeUInt32LE(0x04034b50, 0); // signature
    header.writeUInt16LE(20, 4); // version needed
    header.writeUInt16LE(0, 6); // flags
    header.writeUInt16LE(0, 8); // compression (store)
    header.writeUInt16LE(0, 10); // mod time
    header.writeUInt16LE(0, 12); // mod date
    header.writeUInt32LE(crc, 14); // crc32
    header.writeUInt32LE(contentBuffer.length, 18); // compressed size
    header.writeUInt32LE(contentBuffer.length, 22); // uncompressed size
    header.writeUInt16LE(nameBuffer.length, 26); // name length
    header.writeUInt16LE(0, 28); // extra length
    nameBuffer.copy(header, 30);

    // Central directory entry
    const cdEntry = Buffer.alloc(46 + nameBuffer.length);
    cdEntry.writeUInt32LE(0x02014b50, 0); // signature
    cdEntry.writeUInt16LE(20, 4); // version made by
    cdEntry.writeUInt16LE(20, 6); // version needed
    cdEntry.writeUInt16LE(0, 8); // flags
    cdEntry.writeUInt16LE(0, 10); // compression
    cdEntry.writeUInt16LE(0, 12); // mod time
    cdEntry.writeUInt16LE(0, 14); // mod date
    cdEntry.writeUInt32LE(crc, 16); // crc32
    cdEntry.writeUInt32LE(contentBuffer.length, 20); // compressed
    cdEntry.writeUInt32LE(contentBuffer.length, 24); // uncompressed
    cdEntry.writeUInt16LE(nameBuffer.length, 28); // name length
    cdEntry.writeUInt16LE(0, 30); // extra length
    cdEntry.writeUInt16LE(0, 32); // comment length
    cdEntry.writeUInt16LE(0, 34); // disk start
    cdEntry.writeUInt16LE(0, 36); // internal attrs
    cdEntry.writeUInt32LE(0, 38); // external attrs
    cdEntry.writeUInt32LE(this.offset, 42); // offset
    nameBuffer.copy(cdEntry, 46);

    this.files.push(Buffer.concat([header, contentBuffer]));
    this.centralDir.push(cdEntry);
    this.offset += header.length + contentBuffer.length;
  }

  crc32(buf) {
    let crc = 0xFFFFFFFF;
    const table = [];
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let j = 0; j < 8; j++) {
        c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      }
      table[i] = c;
    }
    for (let i = 0; i < buf.length; i++) {
      crc = table[(crc ^ buf[i]) & 0xFF] ^ (crc >>> 8);
    }
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }

  toBuffer() {
    const cdOffset = this.offset;
    const cd = Buffer.concat(this.centralDir);
    
    const eocd = Buffer.alloc(22);
    eocd.writeUInt32LE(0x06054b50, 0);
    eocd.writeUInt16LE(0, 4); // disk
    eocd.writeUInt16LE(0, 6); // disk with cd
    eocd.writeUInt16LE(this.centralDir.length, 8); // entries on disk
    eocd.writeUInt16LE(this.centralDir.length, 10); // total entries
    eocd.writeUInt32LE(cd.length, 12); // cd size
    eocd.writeUInt32LE(cdOffset, 16); // cd offset
    eocd.writeUInt16LE(0, 20); // comment length

    return Buffer.concat([...this.files, cd, eocd]);
  }
}

const zip = new SimpleZip();

// [Content_Types].xml
zip.addFile('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/>
  <Override PartName="/xl/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/>
</Types>`);

// _rels/.rels
zip.addFile('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="xl/webextensions/taskpanes.xml"/>
</Relationships>`);

// xl/workbook.xml
zip.addFile('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Assistant IA" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`);

// xl/_rels/workbook.xml.rels
zip.addFile('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`);

// xl/worksheets/sheet1.xml
zip.addFile('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Assistant IA Excel - Afrique Francophone</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Le panneau de l'assistant devrait s'ouvrir automatiquement.</t></is></c>
    </row>
  </sheetData>
</worksheet>`);

// xl/webextensions/taskpanes.xml - This tells Excel to show the taskpane
zip.addFile('xl/webextensions/taskpanes.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:taskpane dockstate="right" visibility="1" width="350" row="1" locked="0">
    <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
  </wetp:taskpane>
</wetp:taskpanes>`);

// xl/webextensions/_rels/taskpanes.xml.rels
zip.addFile('xl/webextensions/_rels/taskpanes.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="webextension1.xml"/>
</Relationships>`);

// xl/webextensions/webextension1.xml - The key file that references the add-in
const computerName = process.env.COMPUTERNAME || 'DESKTOP-D7109TB';
const sharePath = `\\\\${computerName}\\AssistantIAExcel`;

zip.addFile('xl/webextensions/webextension1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{E5B75A18-6326-4C10-9CE2-78A7A0864B7A}">
  <we:reference id="${ADDIN_ID}" version="1.0.0.0" store="${sharePath}" storeType="FileSystem"/>
  <we:alternateReferences>
    <we:reference id="${ADDIN_ID}" version="1.0.0.0" store="developer" storeType="Registry"/>
  </we:alternateReferences>
  <we:properties>
    <we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot/>
</we:webextension>`);

const outputPath = path.join(process.env.USERPROFILE, 'Desktop', 'AssistantIA.xlsx');
fs.writeFileSync(outputPath, zip.toBuffer());
console.log(`Workbook created: ${outputPath}`);

// Also create a copy in Temp like office-addin-dev-settings does
const tempPath = path.join(process.env.TEMP, `Excel add-in ${ADDIN_ID}.xlsx`);
fs.writeFileSync(tempPath, zip.toBuffer());
console.log(`Temp workbook created: ${tempPath}`);
