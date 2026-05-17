import * as JSZip from 'jszip';
import { IUserProfile } from './UserProfileService';

export class OfficeFileGenerator {

  /**
   * Generate a .docx Word document with user info content
   */
  public async generateDocx(userProfile: IUserProfile, dateStr: string): Promise<Blob> {
    const zip = new JSZip();

    const bodyContent = this._escapeXml([
      `Created by: ${userProfile.displayName}`,
      `Email: ${userProfile.mail || userProfile.userPrincipalName}`,
      `Department: ${userProfile.department}`,
      `Created at: ${dateStr}`,
      '',
      'This is a generated document.'
    ]);

    // [Content_Types].xml
    zip.file('[Content_Types].xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>'
    );

    // _rels/.rels
    zip.file('_rels/.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>'
    );

    // word/_rels/document.xml.rels
    zip.file('word/_rels/document.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '</Relationships>'
    );

    // word/document.xml - paragraphs with content
    const paragraphs = bodyContent.map(function(line: string) {
      return '<w:p><w:r><w:t>' + line + '</w:t></w:r></w:p>';
    }).join('');

    zip.file('word/document.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:body>' +
      paragraphs +
      '</w:body>' +
      '</w:document>'
    );

    return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  }

  /**
   * Generate a .xlsx Excel spreadsheet with user info content
   */
  public async generateXlsx(userProfile: IUserProfile, dateStr: string): Promise<Blob> {
    const zip = new JSZip();

    const rows = [
      ['Field', 'Value'],
      ['Created by', this._escapeXmlStr(userProfile.displayName)],
      ['Email', this._escapeXmlStr(userProfile.mail || userProfile.userPrincipalName)],
      ['Department', this._escapeXmlStr(userProfile.department)],
      ['Created at', this._escapeXmlStr(dateStr)],
      ['', ''],
      ['Note', 'This is a generated document.']
    ];

    // Shared strings
    const allStrings: string[] = [];
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < rows[r].length; c++) {
        allStrings.push(rows[r][c]);
      }
    }

    let sharedStringsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + allStrings.length + '" uniqueCount="' + allStrings.length + '">';
    for (let i = 0; i < allStrings.length; i++) {
      sharedStringsXml += '<si><t>' + allStrings[i] + '</t></si>';
    }
    sharedStringsXml += '</sst>';

    // Sheet data
    let sheetDataXml = '';
    let strIndex = 0;
    for (let ri = 0; ri < rows.length; ri++) {
      const rowNum = ri + 1;
      sheetDataXml += '<row r="' + rowNum + '">';
      for (let ci = 0; ci < rows[ri].length; ci++) {
        const colLetter = ci === 0 ? 'A' : 'B';
        sheetDataXml += '<c r="' + colLetter + rowNum + '" t="s"><v>' + strIndex + '</v></c>';
        strIndex++;
      }
      sheetDataXml += '</row>';
    }

    // [Content_Types].xml
    zip.file('[Content_Types].xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
      '</Types>'
    );

    // _rels/.rels
    zip.file('_rels/.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>'
    );

    // xl/_rels/workbook.xml.rels
    zip.file('xl/_rels/workbook.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' +
      '</Relationships>'
    );

    // xl/workbook.xml
    zip.file('xl/workbook.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>' +
      '</workbook>'
    );

    // xl/worksheets/sheet1.xml
    zip.file('xl/worksheets/sheet1.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetData>' + sheetDataXml + '</sheetData>' +
      '</worksheet>'
    );

    // xl/sharedStrings.xml
    zip.file('xl/sharedStrings.xml', sharedStringsXml);

    return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  /**
   * Generate a .pptx PowerPoint presentation with user info content
   */
  public async generatePptx(userProfile: IUserProfile, dateStr: string): Promise<Blob> {
    const zip = new JSZip();

    const lines = this._escapeXml([
      'Created by: ' + userProfile.displayName,
      'Email: ' + (userProfile.mail || userProfile.userPrincipalName),
      'Department: ' + userProfile.department,
      'Created at: ' + dateStr
    ]);

    // Build text body paragraphs in EMU coordinates
    let bodyParagraphs = '';
    for (let i = 0; i < lines.length; i++) {
      bodyParagraphs +=
        '<a:p><a:r><a:rPr lang="en-US" sz="1800" dirty="0"/><a:t>' + lines[i] + '</a:t></a:r></a:p>';
    }

    // [Content_Types].xml
    zip.file('[Content_Types].xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
      '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' +
      '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>' +
      '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>' +
      '</Types>'
    );

    // _rels/.rels
    zip.file('_rels/.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>' +
      '</Relationships>'
    );

    // ppt/_rels/presentation.xml.rels
    zip.file('ppt/_rels/presentation.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>' +
      '</Relationships>'
    );

    // ppt/presentation.xml
    zip.file('ppt/presentation.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
      '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>' +
      '<p:sldSz cx="12192000" cy="6858000"/>' +
      '<p:notesSz cx="6858000" cy="9144000"/>' +
      '</p:presentation>'
    );

    // ppt/slideMasters/_rels/slideMaster1.xml.rels
    zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
      '</Relationships>'
    );

    // ppt/slideMasters/slideMaster1.xml
    zip.file('ppt/slideMasters/slideMaster1.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>' +
      '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' +
      '</p:sldMaster>'
    );

    // ppt/slideLayouts/_rels/slideLayout1.xml.rels
    zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>' +
      '</Relationships>'
    );

    // ppt/slideLayouts/slideLayout1.xml
    zip.file('ppt/slideLayouts/slideLayout1.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" type="blank">' +
      '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>' +
      '</p:sldLayout>'
    );

    // ppt/slides/_rels/slide1.xml.rels
    zip.file('ppt/slides/_rels/slide1.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
      '</Relationships>'
    );

    // ppt/slides/slide1.xml - the actual slide with title and content
    zip.file('ppt/slides/slide1.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<p:cSld>' +
      '<p:spTree>' +
      '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr/>' +
      // Title shape
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
      '<p:txBody><a:bodyPr/><a:lstStyle/>' +
      '<a:p><a:r><a:rPr lang="en-US" sz="2800" b="1" dirty="0"/><a:t>Generated Document</a:t></a:r></a:p>' +
      '</p:txBody></p:sp>' +
      // Content shape
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="8229600" cy="4525963"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
      '<p:txBody><a:bodyPr/><a:lstStyle/>' +
      bodyParagraphs +
      '</p:txBody></p:sp>' +
      '</p:spTree>' +
      '</p:cSld>' +
      '</p:sld>'
    );

    return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
  }

  /**
   * Generate a plain text file
   */
  public generateTxt(userProfile: IUserProfile, dateStr: string): Blob {
    const content = [
      'Created by: ' + userProfile.displayName,
      'Email: ' + (userProfile.mail || userProfile.userPrincipalName),
      'Department: ' + userProfile.department,
      'Created at: ' + dateStr,
      '',
      'This is a generated document.'
    ].join('\r\n');

    return new Blob([content], { type: 'text/plain;charset=utf-8' });
  }

  private _escapeXml(lines: string[]): string[] {
    const result: string[] = [];
    for (let i = 0; i < lines.length; i++) {
      result.push(this._escapeXmlStr(lines[i]));
    }
    return result;
  }

  private _escapeXmlStr(str: string): string {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }
}
