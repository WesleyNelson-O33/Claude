// CTS brand, read straight off the September Business Review deck's theme.
const THEME = {
  name: 'CTS Seamless AV',
  headFontFace: 'Calibri',
  bodyFontFace: 'Calibri',
  colors: {
    dk1: '242C39', lt1: 'FFFFFF', dk2: '122B82', lt2: 'E3E6EA',
    accent1: '005358', accent2: '3FD0C9', accent3: '6D71FF',
    accent4: '8F4AFA', accent5: 'F33844', accent6: 'A8052E',
    hlink: '122B82', folHlink: '005358',
  },
};
const C = {
  navy: '242C39', white: 'FFFFFF', mint: 'E8F9F8', tile: 'F4F4F4',
  rule: 'E3E6EA', grey: 'C9CED6', eyebrow: '122B82', body: '3A4453',
  teal: '005358', tealLt: '3FD0C9', blue: '6D71FF', purple: '8F4AFA',
  red: 'F33844', bad: 'A8052E', good: '005358', muted: '6B7585',
};
// Department colours follow the pack: Support reds, Production teals, Consulting blues.
const DEPT = {
  Support: C.red, Events: C.teal, Video: C.tealLt,
  Integration: C.blue, Consulting: C.purple, Admin: C.navy, CTS: C.navy,
};
module.exports = { THEME, C, DEPT };
