/*
 * Builds one standalone training manual from a JSON spec.
 *   node tools/build_manual.js manuals/specs/part2.json
 */
const fs = require('fs');
const path = require('path');
const { Document, Packer } = require('docx');
const lib = require('./manual_lib');

const specPath = process.argv[2];
if (!specPath) {
  console.error('usage: node tools/build_manual.js <spec.json>');
  process.exit(1);
}

const root = path.join(__dirname, '..');
const spec = JSON.parse(fs.readFileSync(specPath, 'utf8'));

const children = [];
if (spec.toc) children.push(...lib.contentsPage(spec.title));
children.push(...lib.renderSpec(spec, { nested: false, root }));

const doc = new Document({
  creator: 'Corporate Technology Services',
  title: spec.title,
  description: spec.subtitle,
  numbering: lib.BULLET_NUMBERING,
  styles: lib.DEFAULT_STYLES,
  sections: [lib.sectionShell(spec.runningHead, children)],
});

const out = path.join(root, spec.output);
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(out, buf);
  console.log(`wrote ${out} (${(buf.length / 1024).toFixed(0)} KB)`);
});
