import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';

test('reduced runtime contains exactly the inventoried upstream bytes and no excluded components', () => {
  const root = fileURLToPath(new URL('../../outlook-mail/node/runtime/', import.meta.url));
  const inventory = JSON.parse(fs.readFileSync(new URL('../../outlook-mail/node/RUNTIME-INVENTORY.json', import.meta.url), 'utf8'));
  const actual = [];
  function walk(directory) {
    for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
      const full = path.join(directory, entry.name);
      if (entry.isDirectory()) walk(full);
      else {
        assert.ok(entry.isFile(), `runtime entry must be a regular file: ${full}`);
        actual.push(path.relative(root, full).split(path.sep).join('/'));
      }
    }
  }
  walk(root);
  assert.deepEqual(actual.sort(), inventory.files.map(file => file.path).sort());
  for (const file of inventory.files) {
    const bytes = fs.readFileSync(path.join(root, file.path));
    assert.equal(bytes.length, file.bytes, file.path);
    assert.equal(createHash('sha256').update(bytes).digest('hex'), file.sha256, file.path);
  }
  assert.equal(inventory.files.length, inventory.selection.retained_files);
  assert.equal(inventory.files.length + inventory.excludedFiles.length, inventory.selection.upstream_files);
  assert.equal(new Set([...inventory.files, ...inventory.excludedFiles].map(file => file.path)).size, inventory.selection.upstream_files);
  for (const file of inventory.excludedFiles) {
    assert.ok(file.reason, `missing exclusion reason: ${file.path}`);
    assert.ok(!fs.existsSync(path.join(root, file.path)), `excluded component restored: ${file.path}`);
    assert.doesNotMatch(path.basename(file.path), /license|notice/i);
  }
});
