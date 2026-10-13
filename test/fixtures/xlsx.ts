import { ooxmlPackage } from '../support/ooxml.ts';
import type { Sample } from '../support/samples.ts';

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export const samples: Sample[] = [
  {
    fileName: 'sample.xlsx',
    build: () =>
      ooxmlPackage(
        {
          name: 'xl/workbook.xml',
          contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
          xml: `<workbook xmlns="${MAIN}" xmlns:r="${R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
        },
        [
          {
            name: 'xl/_rels/workbook.xml.rels',
            xml: `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
          },
          {
            name: 'xl/worksheets/sheet1.xml',
            contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
            xml: `<worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Document Lens sample</t></is></c><c r="B1"><v>42</v></c></row></sheetData></worksheet>`,
          },
        ],
      ),
  },
];
