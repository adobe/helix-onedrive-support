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

'use strict';

const assert = require('assert');

const OneDriveMock = require('../src/OneDriveMock.js');
const StatusCodeError = require('../src/StatusCodeError.js');
const exampleBook = require('./fixtures/book.js');

describe('Worksheet Tests', () => {
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
    assert.equal(name, 'sheet');
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
    assert.equal(index, -1);
  });
  it('Get table names', async () => {
    const values = await sheet.getTableNames();
    assert.deepStrictEqual(values, ['table']);
  });
  it('Get used range address', async () => {
    const range = sheet.usedRange();
    const address = await range.getAddress();
    assert.equal(address, 'Sheet1!A1:B4');
  });
  it('Get used range address local', async () => {
    const range = sheet.usedRange();
    const address = await range.getAddressLocal();
    assert.equal(address, 'A1:B4');
  });
  it('Get all data', async () => {
    const range = sheet.usedRange();
    const address = await range.getData();
    assert.deepStrictEqual(address, oneDrive.workbooks[0].data.sheets[0].usedRange);
  });
  it('Get column names', async () => {
    const range = sheet.usedRange();
    const names = await range.getColumnNames();
    assert.deepStrictEqual(names, ['project', '  c r e a t e d  ']);
  });
  it('Get rows as objects', async () => {
    const range = sheet.usedRange();
    const data = await range.getRowsAsObjects();
    assert.deepStrictEqual(data, [
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
});
