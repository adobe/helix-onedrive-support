/*
 * Copyright 2020 Adobe. All rights reserved.
 * This file is licensed to you under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License. You may obtain a copy
 * of the License at http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under
 * the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
 * OF ANY KIND, either express or implied. See the License for the specific language
 * governing permissions and limitations under the License.
 */

/* eslint-env mocha */
import assert from 'assert';
import { OneDriveMock, StatusCodeError } from '../src/index.js';
import exampleBook from './fixtures/book.js';

describe('Worksheet Tests', () => {
  /**
   * @type {import('../src/index.js').Worksheet}
   */
  let sheet;
  let oneDrive;
  let book;
  beforeEach(() => {
    oneDrive = new OneDriveMock()
      .registerWorkbook('my-drive', 'my-item', exampleBook);
    book = oneDrive.getWorkbook();
    sheet = book.worksheet('sheet');
  });

  it('Get the sheet data', async () => {
    const { name } = await sheet.getData();
    assert.strictEqual(name, 'sheet');
  });

  it('Get a sheet that does not exist fails.', async () => {
    sheet = book.worksheet('sheet-not-exist');
    await assert.rejects(async () => sheet.getData(), new StatusCodeError('sheet-not-exist', 404));
  });

  it('Get named items', async () => {
    const values = await sheet.getNamedItems();
    assert.deepStrictEqual(values, [{ name: 'alice', value: '$A2', comment: 'none' }]);
  });
  it('Get named item', async () => {
    const name = 'alice';
    const values = await sheet.getNamedItem(name);
    assert.deepStrictEqual(values, { name: 'alice', value: '$A2', comment: 'none' });
  });
  it('Add named item', async () => {
    const namedItem = { name: 'bob', value: '$B2', comment: 'none' };
    await sheet.addNamedItem(namedItem.name, namedItem.value, namedItem.comment);
    assert.deepStrictEqual(namedItem, {
      comment: 'none',
      name: 'bob',
      value: '$B2',
    });
  });
  it('Delete named item', async () => {
    const name = 'alice';
    await sheet.deleteNamedItem(name);
    const index = oneDrive.workbooks[0].data.sheets[0].namedItems
      .findIndex((item) => item.name === name);
    assert.strictEqual(index, -1);
  });
  it('Get table names', async () => {
    const values = await sheet.getTableNames();
    assert.deepStrictEqual(values, ['table']);
  });
  it('Add table with a generated name', async () => {
    const table = await sheet.addTable('A1:B4', true);
    assert.strictEqual(table.name, 'Table2');
  });
  it('Add table with a specific name', async () => {
    const table = await sheet.addTable('A1:B4', true, 'index_table');
    assert.strictEqual(table.name, 'index_table');
  });
  it('Add table with a specific name that is also the generated one', async () => {
    const table = await sheet.addTable('A1:B4', true, 'Table2');
    assert.strictEqual(table.name, 'Table2');
  });
  it('Add table with an existing name', async () => {
    await assert.rejects(async () => sheet.addTable('A1:B4', true, 'table'), /Table name already exists/);
  });
  it('Get used range address', async () => {
    const range = sheet.usedRange();
    const address = await range.getAddress();
    assert.strictEqual(address, 'sheet!A1:B4');
  });
  it('Get used range address local', async () => {
    const range = sheet.usedRange();
    const address = await range.getAddressLocal();
    assert.strictEqual(address, 'A1:B4');
  });
  it('Get range', async () => {
    const range = sheet.range('A1:B4');
    const address = await range.getAddress();
    assert.strictEqual(address, 'A1:B4');
  });
  it('Get all data', async () => {
    const range = sheet.usedRange();
    const data = await range.getData();
    assert.deepStrictEqual(data, oneDrive.workbooks[0].data.sheets[0].usedRange);
  });
  it('Get column names', async () => {
    const range = sheet.usedRange();
    const names = await range.getColumnNames();
    assert.deepStrictEqual(names, ['project', '  c r e a t e d  ']);
  });
  it('Get rows as objects', async () => {
    const range = sheet.usedRange();
    const data = await range.getData();
    assert.strictEqual(data.address, 'sheet!A1:B4');
    const rows = await range.getRowsAsObjects();
    assert.deepStrictEqual(rows, [
      { '  c r e a t e d  ': 2018, project: 'Helix' },
      { '  c r e a t e d  ': 2019, project: 'What' },
      { '  c r e a t e d  ': 2020, project: 'this' },
      { '  c r e a t e d  ': '\t 2021 ', project: ' Space\u200B ' },
    ]);
  });
  it('Get rows as objects (trimmed)', async () => {
    const range = sheet.usedRange();
    const data = await range.getRowsAsObjects({ trim: true });
    assert.deepStrictEqual(data, [
      { 'c r e a t e d': '2018', project: 'Helix' },
      { 'c r e a t e d': '2019', project: 'What' },
      { 'c r e a t e d': '2020', project: 'this' },
      { 'c r e a t e d': '2021', project: 'Space' },
    ]);
  });
  it('Replace usedRange', async () => {
    const values = [
      ['', ''],
      ['', ''],
      ['', ''],
      ['', ''],
      ['', ''],
    ];
    const range = sheet.usedRange();
    await range.update(values);
    assert.deepStrictEqual(values, await range.getValues());
  });
});
