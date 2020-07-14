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

const MockOneDrive = require('./MockOneDrive.js');
const exampleBook = require('./fixtures/book.js');



const sampleSheet = {
  name: 'sheet',
  tableNames: [
    ['table'],
  ],
  namedItems: [
    { name: 'alice', value: '$A2', comment: 'none' },
  ],
  usedRange: {
    address: 'Sheet1!A1:B4',
    addressLocal: 'A1:B4',
    values: [
      ['project', 'created'],
      ['Helix', 2018],
      ['What', 2019],
      ['this', 2020]],
  },
  ops: ({
    component, command, method, body,
  }) => {
    switch (component) {
      case 'names':
        return namedItemOps(sampleSheet.namedItems)({ command, method, body });
      case 'tables':
        return { value: sampleSheet.tableNames.map((name) => ({ name })) };
      case 'usedRange':
        return sampleSheet.usedRange;
      default:
        return { values: sampleSheet.name };
    }
  },
};

describe('Worksheet Tests', () => {
  let sheet;
  beforeEach(() => {
    const oneDrive = new MockOneDrive()
      .registerWorkbook('my-drive', 'my-item', exampleBook);
    const book = oneDrive.getWorkbook();
    sheet = book.worksheet('sheet');
  });

  it('Get named items', async () => {
    const values = await sheet.getNamedItems();
    assert.deepEqual(values, sampleSheet.namedItems);
  });
  it('Get named item', async () => {
    const name = 'alice';
    const values = await sheet.getNamedItem(name);
    assert.deepEqual(values, sampleSheet.namedItems[0]);
  });
  it('Add named item', async () => {
    const namedItem = { name: 'bob', value: '$B2', comment: 'none' };
    await sheet.addNamedItem(namedItem.name, namedItem.value, namedItem.comment);
    assert.deepEqual(namedItem, sampleSheet.namedItems[sampleSheet.namedItems.length - 1]);
  });
  it('Delete named item', async () => {
    const name = 'alice';
    await sheet.deleteNamedItem(name);
    const index = sampleSheet.namedItems.findIndex((item) => item.name === name);
    assert.equal(index, -1);
  });
  it('Get table names', async () => {
    const values = await sheet.getTableNames();
    assert.deepEqual(values, sampleSheet.tableNames);
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
    assert.deepEqual(address, sampleSheet.usedRange);
  });
  it('Get column names', async () => {
    const range = sheet.usedRange();
    const names = await range.getColumnNames();
    assert.deepEqual(names, ['project', 'created']);
  });
  it('Get rows as objects', async () => {
    const range = sheet.usedRange();
    const data = await range.getRowsAsObjects();
    assert.deepEqual(data, [
      { created: 2018, project: 'Helix' },
      { created: 2019, project: 'What' },
      { created: 2020, project: 'this' },
    ]);
  });
});
