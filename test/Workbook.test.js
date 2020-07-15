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

const TEST_SHARE_LINK = 'https://adobe.sharepoint.com/:x:/r/sites/cg-helix/Shared%20Documents/data-embed-unit-tests/example-data.xlsx';

describe('Workbook Tests', () => {
  let book;
  let sampleBook;
  let oneDrive;
  beforeEach(() => {
    oneDrive = new OneDriveMock()
      .registerWorkbook('my-drive', 'my-item', exampleBook)
      .registerShareLink(TEST_SHARE_LINK, 'my-drive', 'my-item');
    book = oneDrive.getWorkbook();
    sampleBook = oneDrive.workbooks[0].data;
  });

  it('Get workbook via sharelink', async () => {
    const item = await oneDrive.getDriveItemFromShareLink(TEST_SHARE_LINK);
    const workbook = oneDrive.getWorkbook(item);
    assert.equal(workbook.uri, '/drives/my-drive/items/my-item/workbook');
  });

  it('Get workbook via invalid sharelink failed', async () => {
    await assert.rejects(async () => oneDrive.getDriveItemFromShareLink('/not-found'), new StatusCodeError('/not-found', 404));
  });

  it('Get non registered workbook fails', async () => {
    book = oneDrive.getWorkbook({
      id: 'foo',
      parentReference: {
        driveId: 'bar',
      },
    });
    await assert.rejects(async () => book.getWorksheetNames(), new StatusCodeError('', 500));
  });

  it('Get the workbook data', async () => {
    const { name } = await book.getData();
    assert.equal(name, 'book');
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
    assert.deepEqual(sampleBook.namedItems[sampleBook.namedItems.length - 1], namedItem);
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
