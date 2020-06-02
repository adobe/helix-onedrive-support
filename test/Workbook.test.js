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
const Workbook = require('../src/Workbook.js');

const getClient = require('./getClient.js');
const namedItemOps = require('./NamedItemOps.js');

const sampleBook = {
  name: 'book',
  tableNames: [
    ['table'],
  ],
  sheetNames: [
    ['sheet'],
  ],
  namedItems: [
    { name: 'alice', value: '$A2', comment: 'none' },
  ],
  ops: ({
    component, command, method, body,
  }) => {
    switch (component) {
      case 'worksheets':
        return { value: sampleBook.sheetNames.map((name) => ({ name })) };
      case 'tables':
        return { value: sampleBook.tableNames.map((name) => ({ name })) };
      case 'names':
        return namedItemOps(sampleBook.namedItems)({ command, method, body });
      default:
        return { values: sampleBook.name };
    }
  },
};

const oneDrive = {
  getClient: async () => getClient(sampleBook.ops),
};

describe('Workbook Tests', () => {
  const book = new Workbook(oneDrive, '/book', console);
  it('Get sheet names', async () => {
    const values = await book.getWorksheetNames();
    assert.deepEqual(values, sampleBook.sheetNames);
  });
  it('Get table names', async () => {
    const values = await book.getTableNames();
    assert.deepEqual(values, sampleBook.tableNames);
  });
  it('Get named items', async () => {
    const values = await book.getNamedItems();
    assert.deepEqual(values, sampleBook.namedItems);
  });
  it('Get named item', async () => {
    const name = 'alice';
    const values = await book.getNamedItem(name);
    assert.deepEqual(values, sampleBook.namedItems[0]);
  });
  it('Get named item that doesn\'t exist', async () => {
    const name = 'fred';
    const values = await book.getNamedItem(name);
    assert.equal(values, null);
  });
  it('Add named item', async () => {
    const namedItem = { name: 'bob', value: '$B2', comment: 'none' };
    await book.addNamedItem(namedItem.name, namedItem.value, namedItem.comment);
    assert.deepEqual(namedItem, sampleBook.namedItems[sampleBook.namedItems.length - 1]);
  });
  it('Add named item that already exists', async () => {
    const item = { name: 'alice', value: '$B2', comment: 'none' };
    await assert.rejects(async () => book.addNamedItem(item.name, item.value, item.comment), /Named item already exists/);
  });
  it('Delete named item', async () => {
    const name = 'alice';
    await book.deleteNamedItem(name);
    const index = sampleBook.namedItems.findIndex((item) => item.name === name);
    assert.equal(index, -1);
  });
  it('Delete named item that doesn\'t exist', async () => {
    const name = 'fred';
    await assert.rejects(async () => book.deleteNamedItem(name), /Named item not found/);
  });
});
