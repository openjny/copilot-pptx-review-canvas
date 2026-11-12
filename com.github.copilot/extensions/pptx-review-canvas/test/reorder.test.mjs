import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { test } from 'node:test';
import { ViewerError } from '../lib/errors.mjs';
import { findEntry, readEntry, readZipIndex, replaceEntry, validateZip } from '../lib/pptx-zip.mjs';
import { moveOrder, reorderPresentationXml, reorderSlidesInFile } from '../lib/pptx-reorder.mjs';
import { makePptxBuffer, makeTempDir, makeZip, presentationXml } from './helpers.mjs';

const rejects = (fn, code) => assert.throws(fn, (error) => error instanceof ViewerError && error.code === code, `expected ${code}`);
const idsOf = (xml) => [...xml.matchAll(/<p:sldId id="(\d+)"/g)].map((match) => match[1]);
const sectionsOf = (xml) => [...xml.matchAll(/<p14:section name="([^"]*)"[^>]*>([\s\S]*?)<\/p14:section>/g)].map((match) => [match[1], [...match[2].matchAll(/<p14:sldId id="(\d+)"/g)].map((item) => item[1])]);

test('readZipIndex lists the entries, and readEntry returns them whether stored, deflated or followed by a data descriptor', () => {
  const buf = makeZip([
    { name: 'a.txt', data: 'alpha', method: 0 },
    { name: 'dir/b.xml', data: '<b>' + 'beta '.repeat(100) + '</b>', method: 8, descriptor: true },
    { name: '日本語.txt', data: 'こんにちは', method: 8 },
  ], { comment: 'a comment that is part of the end record' });
  const index = readZipIndex(buf);
  assert.deepEqual(index.entries.map((entry) => entry.name), ['a.txt', 'dir/b.xml', '日本語.txt']);
  assert.equal(readEntry(buf, findEntry(index, 'a.txt')).toString(), 'alpha');
  assert.match(readEntry(buf, findEntry(index, 'dir/b.xml')).toString(), /^<b>beta beta/);
  assert.equal(readEntry(buf, findEntry(index, '日本語.txt')).toString(), 'こんにちは');
  assert.equal(findEntry(index, 'missing'), null);
  assert.equal(validateZip(buf).entries.length, 3);
});

test('the ZIP reader refuses what it cannot handle instead of guessing', () => {
  rejects(() => readZipIndex(Buffer.from('not a zip at all, long enough to hold an end record')), 'invalid_pptx');
  rejects(() => readZipIndex(Buffer.alloc(5)), 'invalid_pptx');
  const good = makeZip([{ name: 'a', data: 'x' }]);
  const cut = good.subarray(0, good.length - 30);
  rejects(() => readZipIndex(cut), 'invalid_pptx');

  const zip64 = Buffer.from(good);
  zip64.writeUInt32LE(0xffffffff, zip64.length - 22 + 16);
  rejects(() => readZipIndex(zip64), 'invalid_pptx');

  const encrypted = Buffer.from(good);
  const index = readZipIndex(good);
  encrypted.writeUInt16LE(encrypted.readUInt16LE(index.entries[0].cenOffset + 8) | 1, index.entries[0].cenOffset + 8);
  rejects(() => readEntry(encrypted, readZipIndex(encrypted).entries[0]), 'invalid_pptx');

  const damaged = Buffer.from(makeZip([{ name: 'a', data: 'some data that gets damaged', method: 0 }]));
  const at = readZipIndex(damaged).entries[0];
  damaged[at.localOffset + 30 + 1 + 4] ^= 0xff;
  rejects(() => readEntry(damaged, readZipIndex(damaged).entries[0]), 'invalid_pptx');
});

test('replaceEntry changes one entry and leaves every other entry intact, wherever it sits and however it ends', () => {
  const entries = [
    { name: 'first.bin', data: Buffer.alloc(300, 7), method: 0 },
    { name: 'ppt/presentation.xml', data: '<old>' + 'z'.repeat(400) + '</old>', method: 8, descriptor: true },
    { name: 'after1.xml', data: '<a>' + 'q'.repeat(200) + '</a>', method: 8, descriptor: true },
    { name: 'after2.bin', data: Buffer.from('plain bytes after'), method: 0 },
  ];
  for (const comment of ['', 'archive comment']) {
    const original = makeZip(entries, { comment });
    const index = readZipIndex(original);
    const target = findEntry(index, 'ppt/presentation.xml');
    for (const replacement of ['<new/>', '<new>' + 'y'.repeat(5000) + '</new>']) {
      const changed = replaceEntry(original, index, target, Buffer.from(replacement));
      const next = validateZip(changed);
      assert.deepEqual(next.entries.map((entry) => entry.name), entries.map((entry) => entry.name));
      assert.equal(readEntry(changed, findEntry(next, 'ppt/presentation.xml')).toString(), replacement);
      for (const entry of entries.filter((entry) => entry.name !== 'ppt/presentation.xml')) {
        assert.deepEqual(readEntry(changed, findEntry(next, entry.name)), Buffer.from(entry.data), entry.name);
      }
      assert.equal(changed.subarray(changed.length - comment.length).toString(), comment, 'the end record keeps its comment');
      assert.ok(!(findEntry(next, 'ppt/presentation.xml').flags & 8), 'the new entry carries its sizes up front');
    }
  }
});

test('moveOrder says which old slide sits at each new position', () => {
  assert.deepEqual(moveOrder(5, 2, 4), [1, 3, 4, 2, 5]);
  assert.deepEqual(moveOrder(5, 4, 2), [1, 4, 2, 3, 5]);
  assert.deepEqual(moveOrder(5, 1, 5), [2, 3, 4, 5, 1]);
  assert.deepEqual(moveOrder(5, 5, 1), [5, 1, 2, 3, 4]);
  assert.deepEqual(moveOrder(3, 2, 2), [1, 2, 3]);
});

test('reorderPresentationXml moves a slide and leaves the rest of the text as it was', () => {
  const xml = presentationXml([256, 257, 258, 259]);
  const { xml: next, order } = reorderPresentationXml(xml, 4, 2);
  assert.deepEqual(idsOf(next), ['256', '259', '257', '258']);
  assert.deepEqual(order, [1, 4, 2, 3]);
  assert.equal(next.replace(/<p:sldIdLst>.*<\/p:sldIdLst>/, ''), xml.replace(/<p:sldIdLst>.*<\/p:sldIdLst>/, ''));
  assert.match(next, /<p:sldId id="259" r:id="rId13"\/>/, 'the relationship of a slide travels with it');
  assert.equal(reorderPresentationXml(xml, 3, 3).xml, xml);

  for (const [from, to] of [[0, 1], [1, 5], [1.5, 2], ['1', 2]]) rejects(() => reorderPresentationXml(xml, from, to), 'slide_out_of_range');
  rejects(() => reorderPresentationXml('<p:presentation/>', 1, 1), 'invalid_pptx');
  rejects(() => reorderPresentationXml(presentationXml([256, 256]), 1, 2), 'unsupported_pptx');
});

test('reorderPresentationXml keeps the sections in step: the moved slide joins the section of its new neighbour', () => {
  const sections = [{ name: 'Intro', ids: [256, 257] }, { name: 'Body', ids: [258, 259, 260] }, { name: 'Empty', ids: [] }, { name: 'End', ids: [261] }];
  const xml = presentationXml([256, 257, 258, 259, 260, 261], { sections });
  const run = (from, to) => {
    const { xml: next } = reorderPresentationXml(xml, from, to);
    const groups = sectionsOf(next);
    assert.deepEqual(groups.flatMap(([, ids]) => ids), idsOf(next), `sections and slides agree after ${from} -> ${to}`);
    return groups;
  };
  assert.deepEqual(run(2, 4), [['Intro', ['256']], ['Body', ['258', '259', '257', '260']], ['Empty', []], ['End', ['261']]]);
  assert.deepEqual(run(4, 1), [['Intro', ['259', '256', '257']], ['Body', ['258', '260']], ['Empty', []], ['End', ['261']]]);
  assert.deepEqual(run(6, 3), [['Intro', ['256', '257', '261']], ['Body', ['258', '259', '260']], ['Empty', []], ['End', []]], 'dropped after slide 2 (the last of Intro), so it joins Intro');
  assert.deepEqual(run(6, 4), [['Intro', ['256', '257']], ['Body', ['258', '261', '259', '260']], ['Empty', []], ['End', []]]);
  assert.deepEqual(run(1, 6), [['Intro', ['257']], ['Body', ['258', '259', '260']], ['Empty', []], ['End', ['261', '256']]]);
  const untouched = reorderPresentationXml(xml, 4, 5).xml;
  assert.ok(untouched.includes('<p14:section name="Intro"'), 'sections that did not change are not rewritten');
});

test('reorderPresentationXml refuses sections it cannot keep consistent, and other namespace prefixes it does not know', () => {
  const mismatched = presentationXml([256, 257, 258], { sections: [{ name: 'A', ids: [257, 256] }, { name: 'B', ids: [258] }] });
  rejects(() => reorderPresentationXml(mismatched, 1, 3), 'unsupported_pptx');
  const unknownPrefix = presentationXml([256, 257]).replace('</p:presentation>', '<p:extLst><p:ext uri="x"><q:sectionLst xmlns:q="urn:x"><q:section/></q:sectionLst></p:ext></p:extLst></p:presentation>');
  rejects(() => reorderPresentationXml(unknownPrefix, 1, 2), 'unsupported_pptx');
  const { xml } = reorderPresentationXml(presentationXml([256, 257, 258], { prefix: 'pr' }), 1, 3);
  assert.deepEqual([...xml.matchAll(/<pr:sldId id="(\d+)"/g)].map((match) => match[1]), ['257', '258', '256']);
});

test('reorderSlidesInFile rewrites only the order, keeps a backup, and leaves the file alone when it cannot', async () => {
  const dir = await makeTempDir();
  const backups = path.join(dir, 'backups');
  const file = path.join(dir, 'deck.pptx');
  const original = makePptxBuffer([256, 257, 258, 259], { sections: [{ name: 'All', ids: [256, 257, 258, 259] }] });
  await fs.writeFile(file, original);

  const { order, changed } = await reorderSlidesInFile(file, 1, 3, { backupDir: backups });
  assert.deepEqual([order, changed], [[2, 3, 1, 4], true]);
  const after = await fs.readFile(file);
  const index = validateZip(after);
  assert.deepEqual(idsOf(readEntry(after, findEntry(index, 'ppt/presentation.xml')).toString()), ['257', '258', '256', '259']);
  const before = readZipIndex(original);
  for (const entry of before.entries.filter((entry) => entry.name !== 'ppt/presentation.xml')) {
    assert.deepEqual(readEntry(after, findEntry(index, entry.name)), readEntry(original, entry), entry.name);
  }
  const kept = await fs.readdir(backups);
  assert.equal(kept.length, 1);
  assert.deepEqual(await fs.readFile(path.join(backups, kept[0])), original, 'the backup is the file as it was');
  assert.deepEqual((await fs.readdir(dir)).filter((name) => name.endsWith('.tmp')), [], 'no temporary file is left');

  assert.equal((await reorderSlidesInFile(file, 2, 2, { backupDir: backups })).changed, false);
  await assert.rejects(() => reorderSlidesInFile(file, 1, 9), (error) => error.code === 'slide_out_of_range');
  await assert.rejects(() => reorderSlidesInFile(path.join(dir, 'old.ppt'), 1, 2), (error) => error.code === 'not_reorderable');
  const text = path.join(dir, 'broken.pptx');
  await fs.writeFile(text, 'this is not a presentation, not even a zip file');
  await assert.rejects(() => reorderSlidesInFile(text, 1, 2), (error) => error.code === 'invalid_pptx');
  const noPart = path.join(dir, 'nopart.pptx');
  await fs.writeFile(noPart, makeZip([{ name: 'readme.txt', data: 'hi' }]));
  await assert.rejects(() => reorderSlidesInFile(noPart, 1, 2), (error) => error.code === 'invalid_pptx');

  await fs.chmod(file, 0o444);
  const locked = await fs.readFile(file);
  await assert.rejects(() => reorderSlidesInFile(file, 1, 2, { backupDir: backups }), (error) => error.code === 'file_locked');
  await fs.chmod(file, 0o666);
  assert.deepEqual(await fs.readFile(file), locked, 'a file that could not be written is unchanged');
  assert.deepEqual((await fs.readdir(dir)).filter((name) => name.endsWith('.tmp')), []);
});
