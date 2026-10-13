// Number formats: SheetJS's SSF plus the built-in table, Japanese calendar codes and section colors.
import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { builtinFormat, formatNumber, formatText, isDateFormat, isoDate, sectionColor, splitSections } from '../../../src/webview/spreadsheet/model/numberFormat.ts';

/** 2024-01-15 (a Monday) in the 1900 date system. */
const JAN_15_2024 = 45306;

describe('number formats', () => {
  it('formats numbers, percentages, fractions and General', () => {
    assert.equal(formatNumber(1234.5, '#,##0.00', false).text, '1,234.50');
    assert.equal(formatNumber(0.1234, '0.00%', false).text, '12.34%');
    assert.equal(formatNumber(1.25, '# ?/?', false).text, '1 1/4');
    assert.equal(formatNumber(0.1 + 0.2, 'General', false).text, '0.3');
    assert.equal(formatNumber(123456789012345, 'General', false).text, '1.23457E+14');
  });

  it('formats dates in the 1900 and 1904 date systems', () => {
    assert.equal(formatNumber(JAN_15_2024, 'yyyy-mm-dd', false).text, '2024-01-15');
    assert.equal(formatNumber(JAN_15_2024 - 1462, 'yyyy-mm-dd', true).text, '2024-01-15');
    assert.equal(formatNumber(JAN_15_2024 + 0.75, 'm/d/yyyy h:mm AM/PM', false).text, '1/15/2024 6:00 PM');
    assert.equal(formatNumber(1.5, '[h]:mm:ss', false).text, '36:00:00');
    assert.equal(formatNumber(JAN_15_2024, 'yyyy.m.d', false).text, '2024.1.15');
    assert.equal(formatNumber(1.5 / 86_400, 'mm:ss.00', false).text, '00:01.50');
  });

  it('formats Japanese era and weekday codes', () => {
    assert.equal(formatNumber(JAN_15_2024, '[$-411]ggge"年"m"月"d"日"', false).text, '令和6年1月15日');
    assert.equal(formatNumber(JAN_15_2024, '[$-411]ge.m.d', false).text, 'R6.1.15');
    assert.equal(formatNumber(JAN_15_2024, '[$-411]gge"年"', false).text, '令6年');
    assert.equal(formatNumber(32_516, '[$-411]ggge"年"m"月"d"日"', false).text, '平成1年1月8日', 'the first day of Heisei');
    assert.equal(formatNumber(32_515, '[$-411]ggge"年"m"月"d"日"', false).text, '昭和64年1月7日', 'the last day of Showa');
    assert.equal(formatNumber(JAN_15_2024, 'yyyy/m/d(aaa)', false).text, '2024/1/15(月)');
    assert.equal(formatNumber(JAN_15_2024 - 1462, 'aaaa', true).text, '月曜日');
    assert.equal(formatNumber(JAN_15_2024, 'yyyy"年"m"月"d"日"', false).text, '2024年1月15日');
  });

  it('knows the built-in formats, Japanese ones and locale-dependent dates', () => {
    assert.equal(builtinFormat(4), '#,##0.00');
    assert.equal(builtinFormat(14, 'en-US'), 'm/d/yyyy');
    assert.equal(builtinFormat(14, 'ja'), 'yyyy/m/d');
    assert.equal(builtinFormat(31), 'yyyy"年"m"月"d"日"');
    assert.equal(builtinFormat(999), 'General');
    assert.ok(isDateFormat(builtinFormat(22)));
    assert.ok(isDateFormat('[$-411]ge.m.d'));
    assert.ok(!isDateFormat('General'));
    assert.ok(!isDateFormat('#,##0'));
    assert.ok(!isDateFormat('[$-411]0.00E+00'));
    assert.ok(!isDateFormat('[Red'));
  });

  it('picks the section color for the value', () => {
    assert.equal(formatNumber(-1234, '#,##0;[Red]-#,##0', false).color, '#ff0000');
    assert.equal(formatNumber(1234, '#,##0;[Red]-#,##0', false).color, undefined);
    assert.equal(sectionColor('[Blue]0;[Red]-0;[Green]0', 0), '#00ff00');
    assert.equal(sectionColor('[Color10]0', 5), '#008000');
    assert.equal(sectionColor('[Red][<0]0;[Blue][>100]0;0', -1), '#ff0000');
    assert.equal(sectionColor('[Red][<0]0;[Blue][>100]0;0', 500), '#0000ff');
    assert.equal(sectionColor('[Red][<0]0;[Blue][>100]0;0', 50), undefined);
  });

  it('splits sections outside quotes, brackets and escapes', () => {
    assert.deepEqual(splitSections('0;"a;b";\\;;@'), ['0', '"a;b"', '\\;', '@']);
  });

  it('formats text with a text section', () => {
    assert.deepEqual(formatText('abc', '0;-0;0;[Blue]"<"@">"'), { text: '<abc>', color: '#0000ff' });
    assert.deepEqual(formatText('abc', '0.00'), { text: 'abc' });
  });

  it('writes dates for the cell bar as ISO dates', () => {
    assert.equal(isoDate(JAN_15_2024, false), '2024-01-15');
    assert.equal(isoDate(JAN_15_2024 + 0.5, false), '2024-01-15 12:00:00');
    assert.equal(isoDate(0.25, false), '06:00:00');
    assert.equal(isoDate(JAN_15_2024 - 1462, true), '2024-01-15');
  });
});
