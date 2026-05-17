export interface IDocumentType {
  key: string;
  text: string;
  extension: string;
  icon: string;
  mimeType: string;
  /** SharePoint Online prefix for opening in Office Online (e.g. :w: for Word) */
  onlinePrefix: string;
}

export const DocumentTypes: IDocumentType[] = [
  {
    key: 'docx',
    text: '\u05DE\u05E1\u05DE\u05DA Word (.docx)',
    extension: '.docx',
    icon: 'WordDocument',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    onlinePrefix: ':w:'
  },
  {
    key: 'xlsx',
    text: '\u05D2\u05D9\u05DC\u05D9\u05D5\u05DF Excel (.xlsx)',
    extension: '.xlsx',
    icon: 'ExcelDocument',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    onlinePrefix: ':x:'
  },
  {
    key: 'pptx',
    text: '\u05DE\u05E6\u05D2\u05EA PowerPoint (.pptx)',
    extension: '.pptx',
    icon: 'PowerPointDocument',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    onlinePrefix: ':p:'
  },
  {
    key: 'txt',
    text: '\u05E7\u05D5\u05D1\u05E5 \u05D8\u05E7\u05E1\u05D8 (.txt)',
    extension: '.txt',
    icon: 'TextDocument',
    mimeType: 'text/plain',
    onlinePrefix: ''
  }
];
