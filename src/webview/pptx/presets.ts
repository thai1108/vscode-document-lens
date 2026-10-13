// Preset shape geometries (prstGeom), written in the guide-formula language that custom geometries (custGeom) use:
//   av:   default adjust values, "name value ..."
//   gd:   guides, "name op arg arg arg; ..." with the DrawingML operators (*/ +- +/ ?: abs at2 cat2 cos max min mod
//         pin sat2 sin sqrt tan val)
//   path: one string per path: optional "fill=none|lighten|lightenLess|darken|darkenLess", "stroke=0", "w=N", "h=N",
//         then commands M x y | L x y | A wR hR stAng swAng | Q x1 y1 x y | C x1 y1 x2 y2 x y | Z
//   rect: text rectangle "l t r b"
// Angles are in 60000ths of a degree; coordinates are guide names or numbers. Shapes not listed here are drawn as
// their bounding rectangle (connectors as a straight line).

export interface PresetSource {
  readonly av?: string;
  readonly gd?: string;
  readonly path: string | readonly string[];
  readonly rect?: string;
}

const RECT_PATH = 'M l t L r t L r b L l b Z';
const ELLIPSE_PATH = 'M l vc A wd2 hd2 cd2 cd4 A wd2 hd2 3cd4 cd4 A wd2 hd2 0 cd4 A wd2 hd2 cd4 cd4 Z';
const ELLIPSE_RECT_GUIDES = 'idx cos wd2 2700000; idy sin hd2 2700000; il +- hc 0 idx; ir +- hc idx 0; it +- vc 0 idy; ib +- vc idy 0';

// Shared by the wedge callouts: where the tip is, and which side of the box the wedge leaves from.
const WEDGE_GUIDES = [
  'dxPos */ w adj1 100000',
  'dyPos */ h adj2 100000',
  'xPos +- hc dxPos 0',
  'yPos +- vc dyPos 0',
  'dx */ dxPos h 1',
  'dy */ dyPos w 1',
  'adx abs dx',
  'ady abs dy',
  'dq +- adx 0 ady',
  'xg1 ?: dxPos 7 2',
  'xg2 ?: dxPos 10 5',
  'x1 */ w xg1 12',
  'x2 */ w xg2 12',
  'yg1 ?: dyPos 7 2',
  'yg2 ?: dyPos 10 5',
  'y1 */ h yg1 12',
  'y2 */ h yg2 12',
  't1 ?: dxPos l xPos',
  'xl ?: dq t1 l',
  't2 ?: dyPos x1 xPos',
  'xt ?: dq x1 t2',
  't3 ?: dxPos xPos r',
  'xr ?: dq t3 r',
  't4 ?: dyPos xPos x1',
  'xb ?: dq x1 t4',
  't5 ?: dxPos y1 yPos',
  'yl ?: dq t5 y1',
  't6 ?: dyPos t yPos',
  'yt ?: dq t t6',
  't7 ?: dxPos yPos y1',
  'yr ?: dq t7 y1',
  't8 ?: dyPos yPos b',
  'yb ?: dq b t8',
].join('; ');

const CALLOUT1_GUIDES = 'y1 */ h adj1 100000; x1 */ w adj2 100000; y2 */ h adj3 100000; x2 */ w adj4 100000';
const CALLOUT1_AV = 'adj1 18750 adj2 -8333 adj3 112500 adj4 -38333';
const CALLOUT2_GUIDES = `${CALLOUT1_GUIDES}; y3 */ h adj5 100000; x3 */ w adj6 100000`;
const CALLOUT2_AV = 'adj1 18750 adj2 -8333 adj3 18750 adj4 -16667 adj5 112500 adj6 -46667';

export const PRESETS: Readonly<Record<string, PresetSource>> = {
  rect: { path: RECT_PATH },
  line: { path: 'fill=none M l t L r b' },
  straightConnector1: { path: 'fill=none M l t L r b' },
  bentConnector2: { path: 'fill=none M l t L r t L r b' },
  bentConnector3: { av: 'adj1 50000', gd: 'x1 */ w adj1 100000', path: 'fill=none M l t L x1 t L x1 b L r b' },
  bentConnector4: {
    av: 'adj1 50000 adj2 50000',
    gd: 'x1 */ w adj1 100000; y2 */ h adj2 100000',
    path: 'fill=none M l t L x1 t L x1 y2 L r y2 L r b',
  },
  bentConnector5: {
    av: 'adj1 50000 adj2 50000 adj3 50000',
    gd: 'x1 */ w adj1 100000; x3 */ w adj3 100000; y2 */ h adj2 100000',
    path: 'fill=none M l t L x1 t L x1 y2 L x3 y2 L x3 b L r b',
  },
  curvedConnector2: { path: 'fill=none M l t C wd2 t r hd2 r b' },
  curvedConnector3: {
    av: 'adj1 50000',
    gd: 'x2 */ w adj1 100000; x1 +/ l x2 2; x3 +/ r x2 2; y3 */ h 3 4',
    path: 'fill=none M l t C x1 t x2 hd4 x2 vc C x2 y3 x3 b r b',
  },
  roundRect: {
    av: 'adj 16667',
    gd: 'a pin 0 adj 50000; x1 */ ss a 100000; x2 +- r 0 x1; y2 +- b 0 x1; il */ x1 29289 100000; ir +- r 0 il; ib +- b 0 il',
    path: 'M l x1 A x1 x1 cd2 cd4 L x2 t A x1 x1 3cd4 cd4 L r y2 A x1 x1 0 cd4 L x1 b A x1 x1 cd4 cd4 Z',
    rect: 'il il ir ib',
  },
  ellipse: { gd: ELLIPSE_RECT_GUIDES, path: ELLIPSE_PATH, rect: 'il it ir ib' },
  triangle: {
    av: 'adj 50000',
    gd: 'a pin 0 adj 100000; x1 */ w a 200000; x2 */ w a 100000; x3 +- x1 wd2 0',
    path: 'M l b L x2 t L r b Z',
    rect: 'x1 vc x3 b',
  },
  rtTriangle: { gd: 'il */ w 1 12; it */ h 7 12; ir */ w 7 12; ib */ h 11 12', path: 'M l t L r b L l b Z', rect: 'il it ir ib' },
  diamond: { gd: 'ir */ w 3 4; ib */ h 3 4', path: 'M l vc L hc t L r vc L hc b Z', rect: 'wd4 hd4 ir ib' },
  parallelogram: {
    av: 'adj 25000',
    gd: 'maxAdj */ 100000 w ss; a pin 0 adj maxAdj; x1 */ ss a 200000; x2 */ ss a 100000; x6 +- r 0 x2; x5 +- r 0 x1; q1 */ 5 a maxAdj; q2 +/ 1 q1 12; il */ q2 w 1; it */ q2 h 1; ir +- r 0 il; ib +- b 0 it',
    path: 'M l b L x2 t L r t L x6 b Z',
    rect: 'il it ir ib',
  },
  trapezoid: {
    av: 'adj 25000',
    gd: 'maxAdj */ 50000 w ss; a pin 0 adj maxAdj; x1 */ ss a 200000; x2 */ ss a 100000; x3 +- r 0 x2; x4 +- r 0 x1; il */ wd3 a maxAdj; it */ hd3 a maxAdj; ir +- r 0 il',
    path: 'M l b L x2 t L x3 t L r b Z',
    rect: 'il it ir b',
  },
  pentagon: {
    av: 'hf 105146 vf 110557',
    gd: 'swd2 */ wd2 hf 100000; shd2 */ hd2 vf 100000; svc */ vc vf 100000; dx1 cos swd2 1080000; dx2 cos swd2 18360000; dy1 sin shd2 1080000; dy2 sin shd2 18360000; x1 +- hc 0 dx1; x2 +- hc 0 dx2; x3 +- hc dx2 0; x4 +- hc dx1 0; y1 +- svc 0 dy1; y2 +- svc 0 dy2; it */ y1 dx2 dx1',
    path: 'M x1 y1 L hc t L x4 y1 L x3 y2 L x2 y2 Z',
    rect: 'x2 it x3 y2',
  },
  hexagon: {
    av: 'adj 25000 vf 115470',
    gd: 'maxAdj */ 50000 w ss; a pin 0 adj maxAdj; shd2 */ hd2 vf 100000; x1 */ ss a 100000; x2 +- r 0 x1; dy1 sin shd2 3600000; y1 +- vc 0 dy1; y2 +- vc dy1 0; q1 */ maxAdj -1 2; q2 +- a q1 0; q3 ?: q2 4 2; q4 ?: q2 3 2; q5 ?: q2 q1 0; q6 +/ a q5 q1; q7 */ q6 q4 -1; q8 +- q3 q7 0; il */ w q8 24; it */ h q8 24; ir +- r 0 il; ib +- b 0 it',
    path: 'M l vc L x1 y1 L x2 y1 L r vc L x2 y2 L x1 y2 Z',
    rect: 'il it ir ib',
  },
  octagon: {
    av: 'adj 29289',
    gd: 'a pin 0 adj 50000; x1 */ ss a 100000; x2 +- r 0 x1; y2 +- b 0 x1; il */ x1 1 2; ir +- r 0 il; ib +- b 0 il',
    path: 'M l x1 L x1 t L x2 t L r x1 L r y2 L x2 b L x1 b L l y2 Z',
    rect: 'il il ir ib',
  },
  plus: {
    av: 'adj 25000',
    gd: 'a pin 0 adj 50000; x1 */ ss a 100000; x2 +- r 0 x1; y2 +- b 0 x1; d +- w 0 h; il ?: d l x1; ir ?: d r x2; it ?: d x1 t; ib ?: d y2 b',
    path: 'M l x1 L x1 x1 L x1 t L x2 t L x2 x1 L r x1 L r y2 L x2 y2 L x2 b L x1 b L x1 y2 L l y2 Z',
    rect: 'il it ir ib',
  },
  chevron: {
    av: 'adj 50000',
    gd: 'maxAdj */ 100000 w ss; a pin 0 adj maxAdj; x1 */ ss a 100000; x2 +- r 0 x1; dx +- x2 0 x1; il ?: dx x1 l; ir ?: dx x2 r',
    path: 'M l t L x2 t L r vc L x2 b L l b L x1 vc Z',
    rect: 'il t ir b',
  },
  homePlate: {
    av: 'adj 50000',
    gd: 'maxAdj */ 100000 w ss; a pin 0 adj maxAdj; dx1 */ ss a 100000; x1 +- r 0 dx1; ir +/ x1 r 2',
    path: 'M l t L x1 t L r vc L x1 b L l b Z',
    rect: 'l t ir b',
  },
  rightArrow: {
    av: 'adj1 50000 adj2 50000',
    gd: 'maxAdj2 */ 100000 w ss; a1 pin 0 adj1 100000; a2 pin 0 adj2 maxAdj2; dx1 */ ss a2 100000; x1 +- r 0 dx1; dy1 */ h a1 200000; y1 +- vc 0 dy1; y2 +- vc dy1 0; dx2 */ y1 dx1 hd2; x2 +- x1 dx2 0',
    path: 'M l y1 L x1 y1 L x1 t L r vc L x1 b L x1 y2 L l y2 Z',
    rect: 'l y1 x2 y2',
  },
  leftArrow: {
    av: 'adj1 50000 adj2 50000',
    gd: 'maxAdj2 */ 100000 w ss; a1 pin 0 adj1 100000; a2 pin 0 adj2 maxAdj2; dx2 */ ss a2 100000; x2 +- l dx2 0; dy1 */ h a1 200000; y1 +- vc 0 dy1; y2 +- vc dy1 0; dx1 */ y1 dx2 hd2; x1 +- x2 0 dx1',
    path: 'M l vc L x2 t L x2 y1 L r y1 L r y2 L x2 y2 L x2 b Z',
    rect: 'x1 y1 r y2',
  },
  upArrow: {
    av: 'adj1 50000 adj2 50000',
    gd: 'maxAdj2 */ 100000 h ss; a1 pin 0 adj1 100000; a2 pin 0 adj2 maxAdj2; dy2 */ ss a2 100000; y2 +- t dy2 0; dx1 */ w a1 200000; x1 +- hc 0 dx1; x2 +- hc dx1 0; dy1 */ x1 dy2 wd2; y1 +- y2 0 dy1',
    path: 'M l y2 L hc t L r y2 L x2 y2 L x2 b L x1 b L x1 y2 Z',
    rect: 'x1 y1 x2 b',
  },
  downArrow: {
    av: 'adj1 50000 adj2 50000',
    gd: 'maxAdj2 */ 100000 h ss; a1 pin 0 adj1 100000; a2 pin 0 adj2 maxAdj2; dy1 */ ss a2 100000; y1 +- b 0 dy1; dx1 */ w a1 200000; x1 +- hc 0 dx1; x2 +- hc dx1 0; dy2 */ x1 dy1 wd2; y2 +- y1 dy2 0',
    path: 'M l y1 L x1 y1 L x1 t L x2 t L x2 y1 L r y1 L hc b Z',
    rect: 'x1 t x2 y2',
  },
  leftRightArrow: {
    av: 'adj1 50000 adj2 50000',
    gd: 'maxAdj2 */ 50000 w ss; a1 pin 0 adj1 100000; a2 pin 0 adj2 maxAdj2; x2 */ ss a2 100000; x3 +- r 0 x2; dy */ h a1 200000; y1 +- vc 0 dy; y2 +- vc dy 0; dx1 */ y1 x2 hd2; x1 +- x2 0 dx1; x4 +- x3 dx1 0',
    path: 'M l vc L x2 t L x2 y1 L x3 y1 L x3 t L r vc L x3 b L x3 y2 L x2 y2 L x2 b Z',
    rect: 'x1 y1 x4 y2',
  },
  upDownArrow: {
    av: 'adj1 50000 adj2 50000',
    gd: 'maxAdj2 */ 50000 h ss; a1 pin 0 adj1 100000; a2 pin 0 adj2 maxAdj2; y2 */ ss a2 100000; y3 +- b 0 y2; dx1 */ w a1 200000; x1 +- hc 0 dx1; x2 +- hc dx1 0; dy1 */ x1 y2 wd2; y1 +- y2 0 dy1; y4 +- y3 dy1 0',
    path: 'M l y2 L hc t L r y2 L x2 y2 L x2 y3 L r y3 L hc b L l y3 L x1 y3 L x1 y2 Z',
    rect: 'x1 y1 x2 y4',
  },
  notchedRightArrow: {
    av: 'adj1 50000 adj2 50000',
    gd: 'maxAdj2 */ 100000 w ss; a1 pin 0 adj1 100000; a2 pin 0 adj2 maxAdj2; dx2 */ ss a2 100000; x2 +- r 0 dx2; dy1 */ h a1 200000; y1 +- vc 0 dy1; y2 +- vc dy1 0; x1 */ dy1 dx2 hd2; x3 +- r 0 x1',
    path: 'M l y1 L x2 y1 L x2 t L r vc L x2 b L x2 y2 L l y2 L x1 vc Z',
    rect: 'x1 y1 x3 y2',
  },
  stripedRightArrow: {
    av: 'adj1 50000 adj2 50000',
    gd: 'maxAdj2 */ 84375 w ss; a1 pin 0 adj1 100000; a2 pin 0 adj2 maxAdj2; x4 */ ss 5 32; dx5 */ ss a2 100000; x5 +- r 0 dx5; dy1 */ h a1 200000; y1 +- vc 0 dy1; y2 +- vc dy1 0; dx6 */ dy1 dx5 hd2; x6 +- r 0 dx6',
    path: 'M l y1 L ssd32 y1 L ssd32 y2 L l y2 Z M ssd16 y1 L ssd8 y1 L ssd8 y2 L ssd16 y2 Z M x4 y1 L x5 y1 L x5 t L r vc L x5 b L x5 y2 L x4 y2 Z',
    rect: 'x4 y1 x6 y2',
  },
  bentArrow: {
    av: 'adj1 25000 adj2 25000 adj3 25000 adj4 43750',
    gd: 'a2 pin 0 adj2 50000; maxAdj1 */ a2 2 1; a1 pin 0 adj1 maxAdj1; a3 pin 0 adj3 50000; th */ ss a1 100000; aw2 */ ss a2 100000; th2 */ th 1 2; dh2 +- aw2 0 th2; ah */ ss a3 100000; bw +- r 0 ah; bh +- b 0 dh2; bs min bw bh; maxAdj4 */ 100000 bs ss; a4 pin 0 adj4 maxAdj4; bd */ ss a4 100000; bd3 +- bd 0 th; bd2 max bd3 0; x3 +- th bd2 0; x4 +- r 0 ah; y3 +- dh2 th 0; y4 +- y3 dh2 0; y5 +- dh2 bd 0; y6 +- y3 bd2 0',
    path: 'M l b L l y5 A bd bd cd2 cd4 L x4 dh2 L x4 t L r aw2 L x4 y4 L x4 y3 L x3 y3 A bd2 bd2 3cd4 -5400000 L th b Z',
  },
  star4: {
    av: 'adj 12500',
    gd: 'a pin 0 adj 50000; iwd2 */ wd2 a 50000; ihd2 */ hd2 a 50000; sdx cos iwd2 2700000; sdy sin ihd2 2700000; sx1 +- hc 0 sdx; sx2 +- hc sdx 0; sy1 +- vc 0 sdy; sy2 +- vc sdy 0',
    path: 'M l vc L sx1 sy1 L hc t L sx2 sy1 L r vc L sx2 sy2 L hc b L sx1 sy2 Z',
    rect: 'sx1 sy1 sx2 sy2',
  },
  star5: {
    av: 'adj 19098 hf 105146 vf 110557',
    gd: 'a pin 0 adj 50000; swd2 */ wd2 hf 100000; shd2 */ hd2 vf 100000; svc */ vc vf 100000; dx1 cos swd2 1080000; dx2 cos swd2 18360000; dy1 sin shd2 1080000; dy2 sin shd2 18360000; x1 +- hc 0 dx1; x2 +- hc 0 dx2; x3 +- hc dx2 0; x4 +- hc dx1 0; y1 +- svc 0 dy1; y2 +- svc 0 dy2; iwd2 */ swd2 a 50000; ihd2 */ shd2 a 50000; sdx1 cos iwd2 20520000; sdx2 cos iwd2 3240000; sdy1 sin ihd2 3240000; sdy2 sin ihd2 20520000; sx1 +- hc 0 sdx1; sx2 +- hc 0 sdx2; sx3 +- hc sdx2 0; sx4 +- hc sdx1 0; sy1 +- svc 0 sdy1; sy2 +- svc 0 sdy2; sy3 +- svc ihd2 0',
    path: 'M x1 y1 L sx2 sy1 L hc t L sx3 sy1 L x4 y1 L sx4 sy2 L x3 y2 L hc sy3 L x2 y2 L sx1 sy2 Z',
    rect: 'sx1 sy1 sx4 sy3',
  },
  wedgeRectCallout: {
    av: 'adj1 -20833 adj2 62500',
    gd: WEDGE_GUIDES,
    path: 'M l t L x1 t L xt yt L x2 t L r t L r y1 L xr yr L r y2 L r b L x2 b L xb yb L x1 b L l b L l y2 L xl yl L l y1 Z',
  },
  wedgeRoundRectCallout: {
    av: 'adj1 -20833 adj2 62500 adj3 16667',
    gd: `${WEDGE_GUIDES}; dr */ ss adj3 100000; rx +- r 0 dr; by +- b 0 dr`,
    path: 'M l dr A dr dr cd2 cd4 L x1 t L xt yt L x2 t L rx t A dr dr 3cd4 cd4 L r y1 L xr yr L r y2 L r by A dr dr 0 cd4 L x2 b L xb yb L x1 b L dr b A dr dr cd4 cd4 L l y2 L xl yl L l y1 Z',
  },
  wedgeEllipseCallout: {
    av: 'adj1 -20833 adj2 62500',
    gd: `dxPos */ w adj1 100000; dyPos */ h adj2 100000; xPos +- hc dxPos 0; yPos +- vc dyPos 0; sdx */ dxPos h 1; sdy */ dyPos w 1; pang at2 sdx sdy; stAng +- pang 660000 0; enAng +- pang 0 660000; dx1 cos wd2 stAng; dy1 sin hd2 stAng; x1 +- hc dx1 0; y1 +- vc dy1 0; dx2 cos wd2 enAng; dy2 sin hd2 enAng; stAng1 at2 dx1 dy1; enAng1 at2 dx2 dy2; swAng1 +- enAng1 0 stAng1; swAng2 +- swAng1 21600000 0; swAng ?: swAng1 swAng1 swAng2; ${ELLIPSE_RECT_GUIDES}`,
    path: 'M xPos yPos L x1 y1 A wd2 hd2 stAng1 swAng Z',
    rect: 'il it ir ib',
  },
  borderCallout1: { av: CALLOUT1_AV, gd: CALLOUT1_GUIDES, path: [RECT_PATH, 'fill=none M x1 y1 L x2 y2'] },
  callout1: { av: CALLOUT1_AV, gd: CALLOUT1_GUIDES, path: [`stroke=0 ${RECT_PATH}`, 'fill=none M x1 y1 L x2 y2'] },
  accentCallout1: { av: CALLOUT1_AV, gd: CALLOUT1_GUIDES, path: [`stroke=0 ${RECT_PATH}`, 'fill=none M x1 t L x1 b', 'fill=none M x1 y1 L x2 y2'] },
  accentBorderCallout1: { av: CALLOUT1_AV, gd: CALLOUT1_GUIDES, path: [RECT_PATH, 'fill=none M x1 t L x1 b', 'fill=none M x1 y1 L x2 y2'] },
  borderCallout2: { av: CALLOUT2_AV, gd: CALLOUT2_GUIDES, path: [RECT_PATH, 'fill=none M x1 y1 L x2 y2 L x3 y3'] },
  callout2: { av: CALLOUT2_AV, gd: CALLOUT2_GUIDES, path: [`stroke=0 ${RECT_PATH}`, 'fill=none M x1 y1 L x2 y2 L x3 y3'] },
  accentCallout2: { av: CALLOUT2_AV, gd: CALLOUT2_GUIDES, path: [`stroke=0 ${RECT_PATH}`, 'fill=none M x1 t L x1 b', 'fill=none M x1 y1 L x2 y2 L x3 y3'] },
  can: {
    av: 'adj 25000',
    gd: 'maxAdj */ 50000 h ss; a pin 0 adj maxAdj; y1 */ ss a 200000; y2 +- y1 y1 0; y3 +- b 0 y1',
    path: [
      'stroke=0 M l y1 A wd2 y1 cd2 -10800000 L r y3 A wd2 y1 0 cd2 Z',
      'stroke=0 fill=lighten M l y1 A wd2 y1 cd2 cd2 A wd2 y1 0 cd2 Z',
      'fill=none M r y1 A wd2 y1 0 cd2 A wd2 y1 cd2 cd2 L r y3 A wd2 y1 0 cd2 L l y1',
    ],
    rect: 'l y2 r y3',
  },
  cube: {
    av: 'adj 25000',
    gd: 'a pin 0 adj 100000; y1 */ ss a 100000; y4 +- b 0 y1; x4 +- r 0 y1',
    path: [
      'stroke=0 M l y1 L x4 y1 L x4 b L l b Z',
      'stroke=0 fill=darkenLess M x4 y1 L r t L r y4 L x4 b Z',
      'stroke=0 fill=lightenLess M l y1 L y1 t L r t L x4 y1 Z',
      'fill=none M l y1 L y1 t L r t L r y4 L x4 b L l b Z M l y1 L x4 y1 L r t M x4 y1 L x4 b',
    ],
    rect: 'l y1 x4 b',
  },
  donut: {
    av: 'adj 25000',
    gd: 'a pin 0 adj 50000; dr */ ss a 100000; iwd2 +- wd2 0 dr; ihd2 +- hd2 0 dr',
    path: `${ELLIPSE_PATH} M dr vc A iwd2 ihd2 cd2 -5400000 A iwd2 ihd2 cd4 -5400000 A iwd2 ihd2 0 -5400000 A iwd2 ihd2 3cd4 -5400000 Z`,
  },
  frame: {
    av: 'adj1 12500',
    gd: 'a1 pin 0 adj1 50000; x1 */ ss a1 100000; x4 +- r 0 x1; y4 +- b 0 x1',
    path: 'M l t L r t L r b L l b Z M x1 x1 L x1 y4 L x4 y4 L x4 x1 Z',
    rect: 'x1 x1 x4 y4',
  },
  foldedCorner: {
    av: 'adj 16667',
    gd: 'a pin 0 adj 50000; dy2 */ ss a 100000; dy1 */ dy2 1 5; x1 +- r 0 dy2; x2 +- x1 dy1 0; y2 +- b 0 dy2; y1 +- y2 dy1 0',
    path: ['stroke=0 M l t L r t L r y2 L x1 b L l b Z', 'stroke=0 fill=darkenLess M x1 b L x2 y1 L r y2 Z', 'fill=none M x1 b L x2 y1 L r y2 M x1 b L l b L l t L r t L r y2'],
    rect: 'l t r y2',
  },
  heart: {
    gd: 'dx1 */ w 49 48; dx2 */ w 10 48; x1 +- hc 0 dx1; x2 +- hc 0 dx2; x3 +- hc dx2 0; x4 +- hc dx1 0; y1 +- t 0 hd3; il */ w 1 6; ir */ w 5 6; ib */ h 2 3',
    path: 'M hc hd4 C x3 y1 x4 hd4 hc b C x1 hd4 x2 y1 hc hd4 Z',
    rect: 'il hd4 ir ib',
  },
  plaque: {
    av: 'adj 16667',
    gd: 'a pin 0 adj 50000; x1 */ ss a 100000; x2 +- r 0 x1; y2 +- b 0 x1; il */ x1 70711 100000; ir +- r 0 il; ib +- b 0 il',
    path: 'M l x1 A x1 x1 cd4 -5400000 L x2 t A x1 x1 cd2 -5400000 L r y2 A x1 x1 3cd4 -5400000 L x1 b A x1 x1 0 -5400000 Z',
    rect: 'il il ir ib',
  },
  teardrop: {
    av: 'adj 100000',
    gd: `a pin 0 adj 200000; r2 sqrt 2; tw */ r2 wd2 1; th */ r2 hd2 1; sw */ tw a 100000; sh */ th a 100000; dx1 cos sw 2700000; dy1 sin sh 2700000; x1 +- hc dx1 0; y1 +- vc 0 dy1; x2 +/ hc x1 2; y2 +/ vc y1 2; ${ELLIPSE_RECT_GUIDES}`,
    path: 'M l vc A wd2 hd2 cd2 cd4 Q x2 t x1 y1 Q r y2 r vc A wd2 hd2 0 cd4 A wd2 hd2 cd4 cd4 Z',
    rect: 'il it ir ib',
  },
  snip1Rect: {
    av: 'adj 16667',
    gd: 'a pin 0 adj 50000; dx1 */ ss a 100000; x1 +- r 0 dx1; it */ dx1 1 2; ir +/ x1 r 2',
    path: 'M l t L x1 t L r dx1 L r b L l b Z',
    rect: 'l it ir b',
  },
  snip2SameRect: {
    av: 'adj1 16667 adj2 0',
    gd: 'a1 pin 0 adj1 50000; a2 pin 0 adj2 50000; tx1 */ ss a1 100000; tx2 +- r 0 tx1; bx1 */ ss a2 100000; bx2 +- r 0 bx1; by1 +- b 0 bx1; d +- tx1 0 bx1; dx ?: d tx1 bx1; il */ dx 1 2; ir +- r 0 il; it */ tx1 1 2; ib +/ by1 b 2',
    path: 'M tx1 t L tx2 t L r tx1 L r by1 L bx2 b L bx1 b L l by1 L l tx1 Z',
    rect: 'il it ir ib',
  },
  round1Rect: {
    av: 'adj 16667',
    gd: 'a pin 0 adj 50000; dx1 */ ss a 100000; x1 +- r 0 dx1; idx */ dx1 29289 100000; ir +- r 0 idx',
    path: 'M l t L x1 t A dx1 dx1 3cd4 cd4 L r b L l b Z',
    rect: 'l t ir b',
  },
  round2SameRect: {
    av: 'adj1 16667 adj2 0',
    gd: 'a1 pin 0 adj1 50000; a2 pin 0 adj2 50000; tx1 */ ss a1 100000; tx2 +- r 0 tx1; bx1 */ ss a2 100000; bx2 +- r 0 bx1; by1 +- b 0 bx1; d +- tx1 0 bx1; tdx */ tx1 29289 100000; bdx */ bx1 29289 100000; il ?: d tdx bdx; ir +- r 0 il; ib +- b 0 bdx',
    path: 'M tx1 t L tx2 t A tx1 tx1 3cd4 cd4 L r by1 A bx1 bx1 0 cd4 L bx1 b A bx1 bx1 cd4 cd4 L l tx1 A tx1 tx1 cd2 cd4 Z',
    rect: 'il tdx ir ib',
  },
  corner: {
    av: 'adj1 50000 adj2 50000',
    gd: 'maxAdj1 */ 100000 h ss; maxAdj2 */ 100000 w ss; a1 pin 0 adj1 maxAdj1; a2 pin 0 adj2 maxAdj2; x1 */ ss a2 100000; dy1 */ ss a1 100000; y1 +- b 0 dy1',
    path: 'M l t L x1 t L x1 y1 L r y1 L r b L l b Z',
  },
  leftBracket: {
    av: 'adj 8333',
    gd: 'maxAdj */ 50000 h ss; a pin 0 adj maxAdj; y1 */ ss a 100000; y2 +- b 0 y1',
    path: ['stroke=0 M r b A w y1 cd4 cd4 L l y1 A w y1 cd2 cd4 Z', 'fill=none M r b A w y1 cd4 cd4 L l y1 A w y1 cd2 cd4'],
  },
  rightBracket: {
    av: 'adj 8333',
    gd: 'maxAdj */ 50000 h ss; a pin 0 adj maxAdj; y1 */ ss a 100000; y2 +- b 0 y1',
    path: ['stroke=0 M l t A w y1 3cd4 cd4 L r y2 A w y1 0 cd4 Z', 'fill=none M l t A w y1 3cd4 cd4 L r y2 A w y1 0 cd4'],
  },
  bracketPair: {
    av: 'adj 16667',
    gd: 'a pin 0 adj 50000; x1 */ ss a 100000; x2 +- r 0 x1; y2 +- b 0 x1; il */ x1 29289 100000; ir +- r 0 il; ib +- b 0 il',
    path: [
      'stroke=0 M l x1 A x1 x1 cd2 cd4 L x2 t A x1 x1 3cd4 cd4 L r y2 A x1 x1 0 cd4 L x1 b A x1 x1 cd4 cd4 Z',
      'fill=none M x1 b A x1 x1 cd4 cd4 L l x1 A x1 x1 cd2 cd4 M x2 t A x1 x1 3cd4 cd4 L r y2 A x1 x1 0 cd4',
    ],
    rect: 'il il ir ib',
  },
  arc: {
    av: 'adj1 16200000 adj2 0',
    gd: 'stAng pin 0 adj1 21599999; enAng pin 0 adj2 21599999; sw11 +- enAng 0 stAng; sw12 +- sw11 21600000 0; swAng ?: sw11 sw11 sw12; wt1 sin wd2 stAng; ht1 cos hd2 stAng; dx1 cat2 wd2 ht1 wt1; dy1 sat2 hd2 ht1 wt1; x1 +- hc dx1 0; y1 +- vc dy1 0',
    path: ['stroke=0 M x1 y1 A wd2 hd2 stAng swAng L hc vc Z', 'fill=none M x1 y1 A wd2 hd2 stAng swAng'],
  },
  pie: {
    av: 'adj1 0 adj2 16200000',
    gd: 'stAng pin 0 adj1 21599999; enAng pin 0 adj2 21599999; sw1 +- enAng 0 stAng; sw2 +- sw1 21600000 0; swAng ?: sw1 sw1 sw2; wt1 sin wd2 stAng; ht1 cos hd2 stAng; dx1 cat2 wd2 ht1 wt1; dy1 sat2 hd2 ht1 wt1; x1 +- hc dx1 0; y1 +- vc dy1 0',
    path: 'M hc vc L x1 y1 A wd2 hd2 stAng swAng Z',
  },
  mathPlus: {
    av: 'adj1 23520',
    gd: 'a1 pin 0 adj1 73490; dx1 */ w 73490 200000; dy1 */ h 73490 200000; dx2 */ ss a1 200000; x1 +- hc 0 dx1; x2 +- hc 0 dx2; x3 +- hc dx2 0; x4 +- hc dx1 0; y1 +- vc 0 dy1; y2 +- vc 0 dx2; y3 +- vc dx2 0; y4 +- vc dy1 0',
    path: 'M x1 y2 L x2 y2 L x2 y1 L x3 y1 L x3 y2 L x4 y2 L x4 y3 L x3 y3 L x3 y4 L x2 y4 L x2 y3 L x1 y3 Z',
    rect: 'x1 y2 x4 y3',
  },
  mathMinus: {
    av: 'adj1 23520',
    gd: 'a1 pin 0 adj1 100000; dy1 */ h a1 200000; dx1 */ w 73490 200000; y1 +- vc 0 dy1; y2 +- vc dy1 0; x1 +- hc 0 dx1; x2 +- hc dx1 0',
    path: 'M x1 y1 L x2 y1 L x2 y2 L x1 y2 Z',
    rect: 'x1 y1 x2 y2',
  },
  flowChartDecision: { gd: 'ir */ w 3 4; ib */ h 3 4', path: 'w=2 h=2 M 0 1 L 1 0 L 2 1 L 1 2 Z', rect: 'wd4 hd4 ir ib' },
  flowChartTerminator: {
    gd: 'il */ w 1018 21600; ir */ w 20582 21600; it */ h 3163 21600; ib */ h 18437 21600',
    path: 'w=21600 h=21600 M 3475 0 L 18125 0 A 3475 10800 3cd4 cd2 L 3475 21600 A 3475 10800 cd4 cd2 Z',
    rect: 'il it ir ib',
  },
  flowChartOffpageConnector: { gd: 'ib */ h 4 5', path: 'w=10 h=10 M 0 0 L 10 0 L 10 8 L 5 10 L 0 8 Z', rect: 'l t r ib' },
  flowChartManualInput: { gd: 'it */ h 1 5', path: 'w=5 h=5 M 0 1 L 5 0 L 5 5 L 0 5 Z', rect: 'l it r b' },
  flowChartPredefinedProcess: {
    gd: 'x2 */ w 1 8; x7 */ w 7 8',
    path: ['w=8 h=8 stroke=0 M 0 0 L 8 0 L 8 8 L 0 8 Z', 'w=8 h=8 fill=none M 1 0 L 1 8 M 7 0 L 7 8', 'w=8 h=8 fill=none M 0 0 L 8 0 L 8 8 L 0 8 Z'],
    rect: 'x2 t x7 b',
  },
  flowChartAlternateProcess: {
    gd: 'x2 +- r 0 ssd6; y2 +- b 0 ssd6; il */ ssd6 29289 100000; ir +- r 0 il; ib +- b 0 il',
    path: 'M l ssd6 A ssd6 ssd6 cd2 cd4 L x2 t A ssd6 ssd6 3cd4 cd4 L r y2 A ssd6 ssd6 0 cd4 L ssd6 b A ssd6 ssd6 cd4 cd4 Z',
    rect: 'il il ir ib',
  },
  flowChartData: { gd: 'x3 */ w 1 5; x4 */ w 4 5', path: 'w=5 h=5 M 0 5 L 1 0 L 5 0 L 4 5 Z', rect: 'x3 t x4 b' },
  flowChartDocument: {
    gd: 'y2 */ h 17322 21600',
    path: 'w=21600 h=21600 M 0 0 L 21600 0 L 21600 17322 C 10800 17322 10800 23922 0 20172 Z',
    rect: 'l t r y2',
  },
  flowChartPreparation: { gd: 'x2 */ w 1 5; x3 */ w 4 5', path: 'w=10 h=10 M 0 5 L 2 0 L 8 0 L 10 5 L 8 10 L 2 10 Z', rect: 'x2 t x3 b' },
  flowChartManualOperation: { gd: 'x3 */ w 1 5; x4 */ w 4 5', path: 'w=5 h=5 M 0 0 L 5 0 L 4 5 L 1 5 Z', rect: 'x3 t x4 b' },
  flowChartMerge: { gd: 'x2 */ w 1 4; x3 */ w 3 4; y2 */ h 1 2', path: 'w=2 h=2 M 0 0 L 2 0 L 1 2 Z', rect: 'x2 t x3 y2' },
  flowChartExtract: { gd: 'x2 */ w 1 4; x3 */ w 3 4; y2 */ h 1 2', path: 'w=2 h=2 M 0 2 L 1 0 L 2 2 Z', rect: 'x2 y2 x3 b' },
  flowChartSort: {
    gd: 'ir */ w 3 4; ib */ h 3 4',
    path: ['w=2 h=2 stroke=0 M 0 1 L 1 0 L 2 1 L 1 2 Z', 'w=2 h=2 fill=none M 0 1 L 2 1', 'w=2 h=2 fill=none M 0 1 L 1 0 L 2 1 L 1 2 Z'],
    rect: 'wd4 hd4 ir ib',
  },
  flowChartCollate: { path: 'w=2 h=2 M 0 0 L 2 0 L 1 1 L 2 2 L 0 2 L 1 1 Z' },
  flowChartPunchedCard: { gd: 'it */ h 1 5', path: 'w=5 h=5 M 0 1 L 1 0 L 5 0 L 5 5 L 0 5 Z', rect: 'l it r b' },
  flowChartDelay: { gd: `${ELLIPSE_RECT_GUIDES}`, path: 'w=2 h=2 M 0 0 L 1 0 A 1 1 3cd4 cd2 L 0 2 Z', rect: 'l it ir ib' },
  flowChartOr: {
    gd: ELLIPSE_RECT_GUIDES,
    path: [`stroke=0 ${ELLIPSE_PATH}`, 'fill=none M hc t L hc b M l vc L r vc', `fill=none ${ELLIPSE_PATH}`],
    rect: 'il it ir ib',
  },
  flowChartSummingJunction: {
    gd: ELLIPSE_RECT_GUIDES,
    path: [`stroke=0 ${ELLIPSE_PATH}`, 'fill=none M il it L ir ib M ir it L il ib', `fill=none ${ELLIPSE_PATH}`],
    rect: 'il it ir ib',
  },
  flowChartInternalStorage: {
    gd: 'x1 */ w 1 8; y1 */ h 1 8',
    path: [`stroke=0 ${RECT_PATH}`, 'fill=none M x1 t L x1 b M l y1 L r y1', `fill=none ${RECT_PATH}`],
    rect: 'x1 y1 r b',
  },
  flowChartMagneticDisk: {
    gd: 'y3 */ h 1 3; y5 */ h 5 6',
    path: ['w=6 h=6 stroke=0 M 0 1 A 3 1 cd2 cd2 L 6 5 A 3 1 0 cd2 Z', 'w=6 h=6 fill=none M 6 1 A 3 1 0 cd2', 'w=6 h=6 fill=none M 0 1 A 3 1 cd2 cd2 L 6 5 A 3 1 0 cd2 Z'],
    rect: 'l y3 r y5',
  },
  flowChartOnlineStorage: { gd: 'x2 */ w 1 6; x5 */ w 5 6', path: 'w=6 h=6 M 1 0 L 6 0 A 1 3 3cd4 -10800000 L 1 6 A 1 3 cd4 cd2 Z', rect: 'x2 t x5 b' },
  flowChartDisplay: { gd: 'x2 */ w 1 6; x5 */ w 5 6', path: 'w=6 h=6 M 0 3 L 1 0 L 5 0 A 1 3 3cd4 cd2 L 1 6 Z', rect: 'x2 t x5 b' },
};

/** Presets drawn with another preset's outline. */
export const PRESET_ALIASES: Readonly<Record<string, string>> = {
  flowChartProcess: 'rect',
  flowChartConnector: 'ellipse',
  flowChartMultidocument: 'flowChartDocument',
  curvedConnector4: 'curvedConnector3',
  curvedConnector5: 'curvedConnector3',
  cloud: 'ellipse',
  cloudCallout: 'wedgeEllipseCallout',
  accentBorderCallout2: 'borderCallout2',
  borderCallout3: 'borderCallout2',
  callout3: 'callout2',
  accentCallout3: 'accentCallout2',
  accentBorderCallout3: 'borderCallout2',
  roundRectCallout: 'wedgeRoundRectCallout',
};

/** Default inner radius (adj, in 1/100000 of the outer radius × 2) of the star presets drawn programmatically. */
export const STAR_POINTS: Readonly<Record<string, readonly [points: number, adj: number]>> = {
  star6: [6, 28868],
  star7: [7, 34601],
  star8: [8, 37500],
  star10: [10, 42533],
  star12: [12, 37500],
  star16: [16, 37500],
  star24: [24, 37500],
  star32: [32, 37500],
};

/** Regular polygons drawn programmatically: number of corners. */
export const POLYGON_CORNERS: Readonly<Record<string, number>> = {
  heptagon: 7,
  decagon: 10,
  dodecagon: 12,
};
