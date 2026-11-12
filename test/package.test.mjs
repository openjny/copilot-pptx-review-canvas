// Checks of the package as a whole: manifests that must agree, the files a catalog needs, the path-length limit, and the links in the documents.
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const readJson = (relative) => JSON.parse(fs.readFileSync(path.join(root, relative), 'utf8'));

const packageJson = readJson('package.json');
const plugin = readJson('plugin.json');
const marketplace = readJson('.github/plugin/marketplace.json');
const extensionDir = 'com.github.copilot/extensions/pptx-review-canvas';
const SEMVER = /^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-[0-9A-Za-z.-]+)?(?:\+[0-9A-Za-z.-]+)?$/;
const MAX_PATH_LENGTH = 72;

function* walk(dir) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    if (['.git', 'node_modules', '.scratch'].includes(entry.name)) continue;
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) yield* walk(full);
    else yield full;
  }
}

test('the manifests name the same plugin at the same version', () => {
  const entry = marketplace.plugins.find((item) => item.name === plugin.name);
  assert.ok(entry, 'the marketplace lists the plugin');
  assert.equal(entry.source, '.');
  assert.equal(packageJson.name, plugin.name);
  for (const version of [packageJson.version, entry.version, marketplace.metadata.version]) assert.equal(version, plugin.version);
  assert.match(plugin.version, SEMVER);
  assert.equal(plugin.license, packageJson.license);
  assert.equal(plugin.$schema, 'https://agent-plugins.org/schemas/1.0.0/plugin.schema.json');
});

// The release-check workflow runs this file on a tag push; elsewhere there is no tag to compare.
test('a release tag is v plus the plugin version', { skip: process.env.GITHUB_REF_TYPE !== 'tag' }, () => {
  assert.equal(process.env.GITHUB_REF_NAME, `v${plugin.version}`);
});

test('the plugin manifest meets what the catalogs ask for', () => {
  assert.match(plugin.name, /^[a-z0-9]([a-z0-9.-]*[a-z0-9])?$/);
  assert.ok(plugin.description.length > 0 && plugin.description.length <= 500);
  assert.ok(plugin.keywords.includes('canvas'));
  assert.ok(plugin.keywords.length <= 10);
  for (const keyword of plugin.keywords) assert.match(keyword, /^[a-z0-9-]{1,30}$/);
  assert.match(plugin.repository, /^https:\/\/github\.com\//);
  assert.equal(plugin.extensions['com.github.copilot'].logo, 'assets/preview.png');
});

test('the preview image is the size catalogs expect', () => {
  const png = fs.readFileSync(path.join(root, 'assets/preview.png'));
  assert.equal(png.subarray(1, 4).toString('latin1'), 'PNG');
  assert.equal(png.readUInt32BE(16), 1280);
  assert.equal(png.readUInt32BE(20), 720);
});

test('the canvas extension is where the plugin format looks for it', () => {
  assert.ok(fs.existsSync(path.join(root, extensionDir, 'extension.mjs')));
  const manifest = readJson(`${extensionDir}/copilot-extension.json`);
  assert.equal(manifest.name, 'pptx-review-canvas');
  assert.equal(manifest.version, Number(plugin.version.split('.')[0]));
});

test(`no file path is longer than ${MAX_PATH_LENGTH} characters`, () => {
  const tooLong = [...walk(root)].map((file) => path.relative(root, file).replaceAll('\\', '/')).filter((relative) => relative.length > MAX_PATH_LENGTH);
  assert.deepEqual(tooLong, []);
});

const MARKDOWN_LINK = /\]\(([^)\s]+)\)|(?:src|href)="([^"]+)"/g;
const withoutCodeFences = (text) => text.replace(/^```[\s\S]*?^```/gm, '');
// GitHub's heading ids: lower case, punctuation removed, spaces to hyphens.
const headingId = (heading) => heading.trim().toLowerCase().replace(/[^\p{L}\p{N}\p{M}_ -]/gu, '').replaceAll(' ', '-');
const headingIds = (file) => new Set([...withoutCodeFences(fs.readFileSync(file, 'utf8')).matchAll(/^#{1,6}[ \t]+(.+?)\s*$/gm)].map((match) => headingId(match[1])));

test('every relative link in the documents leads to a file and a heading that exist', () => {
  const broken = [];
  for (const file of walk(root)) {
    if (!file.endsWith('.md')) continue;
    const document = path.relative(root, file).replaceAll('\\', '/');
    for (const match of withoutCodeFences(fs.readFileSync(file, 'utf8')).matchAll(MARKDOWN_LINK)) {
      const target = match[1] ?? match[2];
      if (/^([a-z][a-z0-9+.-]*:|\/\/)/i.test(target)) continue;
      const [location, anchor] = target.split('#').map(decodeURIComponent);
      const linked = location ? path.resolve(path.dirname(file), location) : file;
      const found = fs.existsSync(linked) && (!anchor || !linked.endsWith('.md') || headingIds(linked).has(anchor));
      if (!found) broken.push(`${document}: ${target}`);
    }
  }
  assert.deepEqual(broken, []);
});
