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

describe('Workbook Tests', () => {
  let book;
  beforeEach(() => {
    const oneDrive = new MockOneDrive()
      .registerWorkbook('my-drive', 'my-item', exampleBook);
    book = oneDrive.getWorkbook();
  });

  it('Get sheet names', async () => {
    const values = await book.getWorksheetNames();
    assert.deepEqual(values, ['sheet']);
  });
  it('Get table names', async () => {
    const values = await book.getTableNames();
    assert.deepEqual(values, ['table']);
  });
  it('Get named items', async () => {
    const values = await book.getNamedItems();
    assert.deepEqual(values, exampleBook.namedItems);
  });
  it('Get named item', async () => {
    const name = 'alice';
    const values = await book.getNamedItem(name);
    assert.deepEqual(values, exampleBook.namedItems[0]);
  });
  it('Get named item that doesn\'t exist', async () => {
    const name = 'fred';
    const values = await book.getNamedItem(name);
    assert.equal(values, null);
  });
  it('Add named item', async () => {
    const namedItem = { name: 'bob', value: '$B2', comment: 'none' };
    await book.addNamedItem(namedItem.name, namedItem.value, namedItem.comment);
    assert.deepEqual(namedItem, exampleBook.namedItems[exampleBook.namedItems.length - 1]);
  });
  it('Add named item that already exists', async () => {
    const item = { name: 'alice', value: '$B2', comment: 'none' };
    await assert.rejects(async () => book.addNamedItem(item.name, item.value, item.comment), /Named item already exists/);
  });
  it('Delete named item', async () => {
    const name = 'alice';
    await book.deleteNamedItem(name);
    const index = exampleBook.namedItems.findIndex((item) => item.name === name);
    assert.equal(index, -1);
  });
  it('Delete named item that doesn\'t exist', async () => {
    const name = 'fred';
    await assert.rejects(async () => book.deleteNamedItem(name), /Named item not found/);
  });
});
