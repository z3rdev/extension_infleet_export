(function initInfleetXlsx(root) {
  "use strict";

  const textEncoder = typeof TextEncoder !== "undefined" ? new TextEncoder() : null;

  function encodeUtf8(value) {
    if (!textEncoder) {
      throw new Error("TextEncoder indisponivel neste navegador.");
    }
    return textEncoder.encode(String(value));
  }

  function escapeXml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function columnName(index) {
    let value = index + 1;
    let name = "";
    while (value > 0) {
      const mod = (value - 1) % 26;
      name = String.fromCharCode(65 + mod) + name;
      value = Math.floor((value - mod) / 26);
    }
    return name;
  }

  function buildSheetXml(rows, columnWidths) {
    const maxColumn = rows.reduce((max, row) => Math.max(max, row.length), 0);
    const dimension = rows.length > 0 && maxColumn > 0
      ? `A1:${columnName(maxColumn - 1)}${rows.length}`
      : "A1";
    const cols = columnWidths
      .map((width, index) => `<col min="${index + 1}" max="${index + 1}" width="${width}" customWidth="1"/>`)
      .join("");
    const sheetData = rows.map((row, rowIndex) => {
      const rowNumber = rowIndex + 1;
      const cells = row.map((value, colIndex) => {
        const ref = `${columnName(colIndex)}${rowNumber}`;
        const style = rowIndex === 0 ? ' s="1"' : "";
        return `<c r="${ref}" t="inlineStr"${style}><is><t>${escapeXml(value)}</t></is></c>`;
      }).join("");
      return `<row r="${rowNumber}">${cells}</row>`;
    }).join("");

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="${dimension}"/>
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <cols>${cols}</cols>
  <sheetData>${sheetData}</sheetData>
  <autoFilter ref="${dimension}"/>
</worksheet>`;
  }

  function buildWorkbookXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Ocorrencias" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;
  }

  function buildStylesXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE8EEF8"/><bgColor indexed="64"/></patternFill></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
  }

  function buildWorkbookRelsXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
  }

  function buildRootRelsXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
  }

  function buildContentTypesXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;
  }

  function makeCrcTable() {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n += 1) {
      let c = n;
      for (let k = 0; k < 8; k += 1) {
        c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      }
      table[n] = c >>> 0;
    }
    return table;
  }

  const crcTable = makeCrcTable();

  function crc32(bytes) {
    let crc = 0xffffffff;
    for (let i = 0; i < bytes.length; i += 1) {
      crc = crcTable[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  function writeUint16(target, offset, value) {
    target[offset] = value & 0xff;
    target[offset + 1] = (value >>> 8) & 0xff;
  }

  function writeUint32(target, offset, value) {
    target[offset] = value & 0xff;
    target[offset + 1] = (value >>> 8) & 0xff;
    target[offset + 2] = (value >>> 16) & 0xff;
    target[offset + 3] = (value >>> 24) & 0xff;
  }

  function concatBytes(parts) {
    const length = parts.reduce((total, part) => total + part.length, 0);
    const output = new Uint8Array(length);
    let offset = 0;
    for (const part of parts) {
      output.set(part, offset);
      offset += part.length;
    }
    return output;
  }

  function zipStore(files) {
    const localParts = [];
    const centralParts = [];
    let offset = 0;
    const flags = 0x0800;
    const method = 0;
    const modTime = 0;
    const modDate = 0x5b21;

    for (const file of files) {
      const nameBytes = encodeUtf8(file.name);
      const dataBytes = encodeUtf8(file.content);
      const crc = crc32(dataBytes);

      const localHeader = new Uint8Array(30 + nameBytes.length);
      writeUint32(localHeader, 0, 0x04034b50);
      writeUint16(localHeader, 4, 20);
      writeUint16(localHeader, 6, flags);
      writeUint16(localHeader, 8, method);
      writeUint16(localHeader, 10, modTime);
      writeUint16(localHeader, 12, modDate);
      writeUint32(localHeader, 14, crc);
      writeUint32(localHeader, 18, dataBytes.length);
      writeUint32(localHeader, 22, dataBytes.length);
      writeUint16(localHeader, 26, nameBytes.length);
      writeUint16(localHeader, 28, 0);
      localHeader.set(nameBytes, 30);
      localParts.push(localHeader, dataBytes);

      const centralHeader = new Uint8Array(46 + nameBytes.length);
      writeUint32(centralHeader, 0, 0x02014b50);
      writeUint16(centralHeader, 4, 20);
      writeUint16(centralHeader, 6, 20);
      writeUint16(centralHeader, 8, flags);
      writeUint16(centralHeader, 10, method);
      writeUint16(centralHeader, 12, modTime);
      writeUint16(centralHeader, 14, modDate);
      writeUint32(centralHeader, 16, crc);
      writeUint32(centralHeader, 20, dataBytes.length);
      writeUint32(centralHeader, 24, dataBytes.length);
      writeUint16(centralHeader, 28, nameBytes.length);
      writeUint16(centralHeader, 30, 0);
      writeUint16(centralHeader, 32, 0);
      writeUint16(centralHeader, 34, 0);
      writeUint16(centralHeader, 36, 0);
      writeUint32(centralHeader, 38, 0);
      writeUint32(centralHeader, 42, offset);
      centralHeader.set(nameBytes, 46);
      centralParts.push(centralHeader);

      offset += localHeader.length + dataBytes.length;
    }

    const centralDirectory = concatBytes(centralParts);
    const end = new Uint8Array(22);
    writeUint32(end, 0, 0x06054b50);
    writeUint16(end, 4, 0);
    writeUint16(end, 6, 0);
    writeUint16(end, 8, files.length);
    writeUint16(end, 10, files.length);
    writeUint32(end, 12, centralDirectory.length);
    writeUint32(end, 16, offset);
    writeUint16(end, 20, 0);

    return concatBytes([...localParts, centralDirectory, end]);
  }

  function buildWorkbook(rows, columns, widths) {
    const tableRows = [
      columns.map((column) => column.label),
      ...rows.map((row) => columns.map((column) => row[column.key] ?? ""))
    ];
    const columnWidths = widths ?? columns.map((column) => column.width ?? 18);
    const files = [
      { name: "[Content_Types].xml", content: buildContentTypesXml() },
      { name: "_rels/.rels", content: buildRootRelsXml() },
      { name: "xl/workbook.xml", content: buildWorkbookXml() },
      { name: "xl/_rels/workbook.xml.rels", content: buildWorkbookRelsXml() },
      { name: "xl/styles.xml", content: buildStylesXml() },
      { name: "xl/worksheets/sheet1.xml", content: buildSheetXml(tableRows, columnWidths) }
    ];

    return zipStore(files);
  }

  const api = { buildWorkbook };
  root.InfleetXlsx = api;

  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
})(typeof globalThis !== "undefined" ? globalThis : window);
