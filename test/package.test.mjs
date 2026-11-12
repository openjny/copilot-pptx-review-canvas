// Checks of the package as a whole: manifests that must agree, the files a catalog needs, and the path-length limit.
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

test('every image the documents link to exists', () => {
  const missing = [];
  for (const document of ['README.md', 'README.ja.md', 'docs/architecture.md']) {
    const file = path.join(root, document);
    if (!fs.existsSync(file)) continue;
    const text = fs.readFileSync(file, 'utf8');
    for (const match of text.matchAll(/(?:\]\(|src=")(\.{1,2}\/[^)"]+\.(?:png|gif|webp))/g)) {
      if (!fs.existsSync(path.resolve(path.dirname(file), match[1]))) missing.push(`${document}: ${match[1]}`);
    }
  }
  assert.deepEqual(missing, []);
});
