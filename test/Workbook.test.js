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
  ops: (component, command, method, body) => {
    let index;
    let len;
    let item;
    switch (component) {
      case 'worksheets':
        return { value: sampleBook.sheetNames.map((name) => ({ name })) };
      case 'tables':
        return { value: sampleBook.tableNames.map((name) => ({ name })) };
      case 'names':
        if (!command) {
          return { value: sampleBook.namedItems };
        }
        if (command === 'add') {
          len = sampleBook.namedItems.push({
            name: body.name,
            value: body.reference,
            comment: body.comment,
          });
          return sampleBook.namedItems[len - 1];
        }
        index = sampleBook.namedItems.findIndex((i) => i.name === command);
        if (index === -1) {
          throw new Error('not found');
        }
        item = sampleBook.namedItems[index];
        if (method === 'DELETE') {
          sampleBook.namedItems.splice(index, 1);
        }
        return item;
      default:
        return { values: sampleBook.name };
    }
  },
};

const oneDrive = {
  getClient: async () => {
    const f = async ({
      uri, method, body,
    }) => {
      const [, , component, command] = uri.split('/');
      return sampleBook.ops(component, command, method, body);
    };
    f.get = (uri) => f({ method: 'GET', uri });
    return f;
  },
};

describe('Worksheet Tests', () => {
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
  it('Add named item', async () => {
    const namedItem = { name: 'bob', value: '$B2', comment: 'none' };
    await book.addNamedItem(namedItem.name, namedItem.value, namedItem.comment);
    assert.deepEqual(namedItem, sampleBook.namedItems[sampleBook.namedItems.length - 1]);
  });
  it('Delete named item', async () => {
    const name = 'alice';
    await book.deleteNamedItem(name);
    const index = sampleBook.namedItems.findIndex((item) => item.name === name);
    assert.equal(index, -1);
  });
});
