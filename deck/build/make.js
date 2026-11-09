const path = require('path');
require('./part0.js').introSlide();        // introduction, department colours
const { pres } = require('./parts.js');    // everything else, snipped from the pack
const { THEME } = require('./theme.js');
const { applyTheme } = require('/root/.claude/skills/synced/e9aa5ee9-a99d-40c9-9bb1-cc5bc5010cf8_723c92f1-3883-471f-8787-0ed0739f050b/pptx/scripts/apply_theme.js');

const OUT = path.join(__dirname, '..', 'CTS Monthly Business Review - September 2026.pptx');
(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log('written:', OUT);
})();
