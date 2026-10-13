// An .xlsx package: workbook, styles, theme and shared strings are read up front; each worksheet only when shown.
import { SHEET_LIMITS } from '../limits.ts';
import { DEFAULT_THEME } from '../model/colors.ts';
import { emptySheet } from '../model/sheet.ts';
import { baseColumnWidthPx, defaultRowHeightPt, maxDigitWidth, pointsToPx } from '../model/sizes.ts';
import { EMPTY_STRINGS, type Sheet, type SheetEntry, type SheetKind, type StringTable, type Workbook } from '../model/types.ts';
import { parseRelationships, relationshipsPartOf, type Relationship } from './relationships.ts';
import { readWorksheet } from './sheet.ts';
import { readSharedStrings } from './strings.ts';
import { parseStyles, parseTheme } from './styles.ts';
import { scanXml } from './xml.ts';
import { ZipReader } from './zip.ts';

const SHEET_KINDS: Record<string, SheetKind> = {
  worksheet: 'worksheet',
  chartsheet: 'chartsheet',
  dialogsheet: 'dialogsheet',
  xlMacrosheet: 'macrosheet',
  xlIntlMacrosheet: 'macrosheet',
};

interface SheetPart extends SheetEntry {
  part?: string;
}

export async function openXlsx(bytes: Uint8Array): Promise<Workbook> {
  const zip = ZipReader.open(bytes);
  const rootRelationships = zip.has('_rels/.rels') ? parseRelationships(await zip.readText('_rels/.rels'), '') : new Map<string, Relationship>();
  const workbookPart = [...rootRelationships.values()].find((relationship) => relationship.type === 'officeDocument' && !relationship.external)?.target ?? 'xl/workbook.xml';
  if (!zip.has(workbookPart)) {
    throw new Error('Not a spreadsheet: the workbook part is missing');
  }
  const relationshipsPart = relationshipsPartOf(workbookPart);
  const relationships = zip.has(relationshipsPart) ? parseRelationships(await zip.readText(relationshipsPart), workbookPart) : new Map<string, Relationship>();
  const partOfType = (type: string) => [...relationships.values()].find((relationship) => relationship.type === type && !relationship.external && zip.has(relationship.target))?.target;

  const notices: string[] = [];
  const { sheets, date1904, activeTab, sheetsCut } = readWorkbookPart(await zip.readText(workbookPart), relationships);
  if (sheetsCut) {
    notices.push(`Only the first ${SHEET_LIMITS.maxSheets} sheets are shown.`);
  }

  const themePart = partOfType('theme');
  const theme = themePart ? parseTheme(await zip.readText(themePart)) : DEFAULT_THEME;
  const stylesPart = partOfType('styles');
  const styles = parseStyles(stylesPart ? await zip.readText(stylesPart) : '', theme);

  let sharedStrings: StringTable = EMPTY_STRINGS;
  const stringsPart = partOfType('sharedStrings');
  if (stringsPart) {
    const result = await readSharedStrings(zip.textChunks(stringsPart), styles.colors);
    sharedStrings = result.table;
    if (result.truncated) {
      notices.push(`Only the first ${SHEET_LIMITS.maxSharedStrings.toLocaleString('en-US')} distinct texts are shown; later ones appear empty.`);
    }
  }

  const visible = sheets.findIndex((sheet) => !sheet.hidden);
  const activeSheet = sheets[activeTab] && !sheets[activeTab].hidden ? activeTab : Math.max(0, visible);

  return {
    sheets: sheets.map(({ name, hidden, kind }) => ({ name, hidden, kind })),
    activeSheet,
    date1904,
    formats: styles.formats,
    sharedStrings,
    notices,
    async loadSheet(index: number, signal?: AbortSignal): Promise<Sheet> {
      const entry = sheets[index];
      if (!entry) {
        throw new Error(`No sheet ${index}`);
      }
      const digitWidth = maxDigitWidth(styles.defaultFont);
      const blank = (notice: string) => emptySheet(entry.name, entry.kind, pointsToPx(defaultRowHeightPt(styles.defaultFont)), baseColumnWidthPx(8, digitWidth), notice);
      if (entry.kind === 'chartsheet') {
        return blank(`“${entry.name}” is a chart sheet; charts are not shown.`);
      }
      if (!entry.part || !zip.has(entry.part)) {
        return blank(`“${entry.name}” is missing from the file.`);
      }
      const sheetRelationshipsPart = relationshipsPartOf(entry.part);
      const sheetRelationships = zip.has(sheetRelationshipsPart) ? parseRelationships(await zip.readText(sheetRelationshipsPart), entry.part) : new Map<string, Relationship>();
      return readWorksheet(
        zip.textChunks(entry.part, signal),
        {
          name: entry.name,
          kind: entry.kind,
          part: entry.part,
          formats: styles.formats,
          sharedStringCount: sharedStrings.texts.length,
          date1904,
          defaultFont: styles.defaultFont,
          colors: styles.colors,
          relationships: sheetRelationships,
        },
        signal,
      );
    },
  };
}

function readWorkbookPart(xml: string, relationships: ReadonlyMap<string, Relationship>): { sheets: SheetPart[]; date1904: boolean; activeTab: number; sheetsCut: boolean } {
  const sheets: SheetPart[] = [];
  let date1904 = false;
  let activeTab = 0;
  let views = 0;
  let sheetsCut = false;
  scanXml(xml, {
    open(tag) {
      switch (tag.name) {
        case 'workbookPr': {
          const value = tag.attr('date1904');
          date1904 = value === '1' || value === 'true';
          return;
        }
        case 'workbookView': {
          views++;
          const tab = Number(tag.attr('activeTab') ?? 0);
          if (views === 1 && Number.isInteger(tab) && tab >= 0) {
            activeTab = tab;
          }
          return;
        }
        case 'sheet': {
          if (sheets.length >= SHEET_LIMITS.maxSheets) {
            sheetsCut = true;
            return;
          }
          const id = tag.attr('id');
          const relationship = id ? relationships.get(id) : undefined;
          const state = tag.attr('state');
          const name = (tag.attr('name') ?? `Sheet${sheets.length + 1}`).slice(0, 255);
          const sheet: SheetPart = {
            name,
            hidden: state === 'hidden' || state === 'veryHidden',
            kind: (relationship && SHEET_KINDS[relationship.type]) ?? 'worksheet',
          };
          if (relationship && !relationship.external) {
            sheet.part = relationship.target;
          }
          sheets.push(sheet);
          return;
        }
      }
    },
  });
  if (!sheets.length) {
    throw new Error('Not a spreadsheet: the workbook lists no sheets');
  }
  return { sheets, date1904, activeTab, sheetsCut };
}
