import assert from 'node:assert/strict';
import { before, describe, it } from 'node:test';
import { fontFamily, resolveTypeface } from '../../../src/webview/pptx/fonts.ts';
import { autoNumber, hasText, readBodyProperties, symbolCharacter } from '../../../src/webview/pptx/text.ts';
import { DEFAULT_THEME, parseTheme } from '../../../src/webview/pptx/theme.ts';
import { A_NS, installDom, xml } from './dom.ts';

describe('numbering and bullets', () => {
  it('formats the auto-numbering schemes', () => {
    assert.equal(autoNumber('arabicPeriod', 3), '3.');
    assert.equal(autoNumber('arabicParenR', 3), '3)');
    assert.equal(autoNumber('arabicParenBoth', 3), '(3)');
    assert.equal(autoNumber('arabicPlain', 12), '12');
    assert.equal(autoNumber('romanUcPeriod', 14), 'XIV.');
    assert.equal(autoNumber('romanLcParenR', 4), 'iv)');
    assert.equal(autoNumber('alphaUcPeriod', 2), 'B.');
    assert.equal(autoNumber('alphaLcParenBoth', 28), '(bb)');
    assert.equal(autoNumber('circleNumDbPlain', 3), '③');
    assert.equal(autoNumber('circleNumWdBlackPlain', 1), '❶');
    assert.equal(autoNumber('ea1JpnChsDbPeriod', 12), '１２．');
    assert.equal(autoNumber('arabic1Minus', 2), '- 2 -');
    assert.equal(autoNumber('somethingNew', 0), '1.');
  });

  it('maps Wingdings and Symbol bullets to Unicode', () => {
    assert.equal(symbolCharacter('Wingdings', '§'), '▪');
    assert.equal(symbolCharacter('Wingdings', String.fromCharCode(0xf0d8)), '➢');
    assert.equal(symbolCharacter('Symbol', '·'), '•');
    assert.equal(symbolCharacter('Wingdings 2', 'l'), '●');
    assert.equal(symbolCharacter('Arial', '•'), undefined);
    assert.equal(symbolCharacter('Wingdings', 'Z'), undefined);
  });
});

describe('text bodies', () => {
  before(() => installDom());

  it('merges body properties from the shape down to its master placeholder', () => {
    const own = xml(`<a:bodyPr ${A_NS} anchor="ctr" lIns="0"><a:normAutofit fontScale="62500" lnSpcReduction="20000"/></a:bodyPr>`);
    const master = xml(`<a:bodyPr ${A_NS} wrap="none" vert="eaVert" rot="5400000" numCol="2" tIns="91440"><a:noAutofit/></a:bodyPr>`);
    const body = readBodyProperties([own, undefined, master]);
    assert.deepEqual(body.insets, { l: 0, t: 9.6, r: 9.6, b: 4.8 });
    assert.equal(body.anchor, 'ctr');
    assert.equal(body.wrap, false);
    assert.equal(body.vert, 'eaVert');
    assert.equal(body.rotation, 90);
    assert.equal(body.columns, 2);
    assert.equal(body.fontScale, 0.625);
    assert.equal(body.lineReduction, 0.2);
    const defaults = readBodyProperties([]);
    assert.deepEqual([defaults.anchor, defaults.wrap, defaults.fontScale, defaults.columns], ['t', true, 1, 1]);
  });

  it('knows when a text body has nothing to show', () => {
    assert.equal(hasText(xml(`<p:txBody xmlns:p="urn:p" ${A_NS}><a:bodyPr/><a:p><a:endParaRPr/></a:p><a:p><a:r><a:t></a:t></a:r></a:p></p:txBody>`)), false);
    assert.equal(hasText(xml(`<p:txBody xmlns:p="urn:p" ${A_NS}><a:p><a:fld type="slidenum"><a:t>‹#›</a:t></a:fld></a:p></p:txBody>`)), true);
    assert.equal(hasText(undefined), false);
  });
});

describe('fonts', () => {
  before(() => installDom());

  it('resolves theme font references, using the script font for East Asian text', () => {
    const theme = parseTheme(
      xml(
        `<a:theme ${A_NS}><a:themeElements><a:fontScheme><a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック"/><a:font script="Hang" typeface="맑은 고딕"/></a:majorFont>` +
          '<a:minorFont><a:latin typeface="Aptos"/><a:ea typeface="Meiryo"/><a:cs typeface="Arial"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>',
      ),
    );
    assert.equal(resolveTypeface('+mj-lt', theme), 'Aptos Display');
    assert.equal(resolveTypeface('+mj-ea', theme, 'ja-JP'), '游ゴシック');
    assert.equal(resolveTypeface('+mj-ea', theme, 'ko-KR'), '맑은 고딕');
    assert.equal(resolveTypeface('+mj-ea', theme, 'en-US'), undefined);
    assert.equal(resolveTypeface('+mn-ea', theme, 'ja-JP'), 'Meiryo');
    assert.equal(resolveTypeface('+mn-cs', theme), 'Arial');
    assert.equal(resolveTypeface('Georgia', theme), 'Georgia');
    assert.equal(resolveTypeface(undefined, theme), undefined);
    assert.equal(resolveTypeface('+mn-lt', DEFAULT_THEME), 'Calibri');
  });

  it('builds a CSS family list with substitutes and a generic family', () => {
    assert.equal(fontFamily('Calibri', undefined), '"Calibri", "Carlito", "Segoe UI", "Helvetica Neue", "Arial", sans-serif');
    assert.match(fontFamily('Times New Roman', 'ＭＳ 明朝'), /^"Times New Roman", "Liberation Serif", "Times", "ＭＳ 明朝", "YuMincho", .*, serif$/);
    assert.match(fontFamily('Consolas'), /, monospace$/);
    assert.equal(fontFamily(undefined, ''), 'sans-serif');
    // Names come from the file and must not break out of the declaration.
    assert.equal(fontFamily('Evil"; } body { color: red'), '"Evil  body  color: red", sans-serif');
  });
});
