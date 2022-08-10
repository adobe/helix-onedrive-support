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

const TEST_SHARE_LINK = 'https://adobe.sharepoint.com/:x:/r/sites/cg-helix/Shared%20Documents/data-embed-unit-tests/example-data.xlsx';

describe('Workbook Tests', () => {
  /**
   * @type {import('../src/index.js').Workbook}
   */
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
    assert.strictEqual(workbook.uri, '/drives/my-drive/items/my-item/workbook');
  });

  it('Get workbook via onedrive uri', async () => {
    const item = await oneDrive.getDriveItemFromShareLink('onedrive:/drives/my-drive/items/my-item');
    const workbook = oneDrive.getWorkbook(item);
    assert.strictEqual(workbook.uri, '/drives/my-drive/items/my-item/workbook');
  });

  it('Get workbook via invalid sharelink failed', async () => {
    await assert.rejects(async () => oneDrive.getDriveItemFromShareLink('https://foo.com/not-found'), new StatusCodeError('https://foo.com/not-found', 404));
  });

  it('Get non registered workbook fails', async () => {
    book = oneDrive.getWorkbook({
      id: 'foo',
      parentReference: {
        driveId: 'bar',
      },
    });
    await assert.rejects(async () => book.getWorksheetNames(), new StatusCodeError('not found', 404));
  });

  it('Get the workbook data', async () => {
    const { name } = await book.getData();
    assert.strictEqual(name, 'book');
  });

  it('Get sheet names', async () => {
    const values = await book.getWorksheetNames();
    assert.deepStrictEqual(values, ['sheet']);
  });
  it('Get table names', async () => {
    const values = await book.getTableNames();
    assert.deepStrictEqual(values, ['table']);
  });
  it('Add table with a generated name', async () => {
    const table = await book.addTable('sheet!A1:B4', true);
    assert.strictEqual(table.name, 'Table2');
  });
  it('Add table with a specific name', async () => {
    const table = await book.addTable('sheet!A1:B4', true, 'index_table');
    assert.strictEqual(table.name, 'index_table');
    const headerNames = await table.getHeaderNames();
    assert.strictEqual(headerNames[0], 'project');
    const row = await table.getRow(0);
    assert.strictEqual(row[0], 'Helix');
  });
  it('Add table with a specific name that is also the generated one', async () => {
    const table = await book.addTable('sheet!A1:B4', false, 'Table2');
    assert.strictEqual(table.name, 'Table2');
  });
  it('Add table with an existing name', async () => {
    await assert.rejects(async () => book.addTable('sheet!A1:B4', true, 'table'), /Table name already exists/);
  });
  it('Get named items', async () => {
    const values = await book.getNamedItems();
    assert.deepStrictEqual(values, sampleBook.namedItems);
  });
  it('Get named item', async () => {
    const name = 'alice';
    const values = await book.getNamedItem(name);
    assert.deepStrictEqual(values, sampleBook.namedItems[0]);
  });
  it('Get named item that doesn\'t exist', async () => {
    const name = 'fred';
    const values = await book.getNamedItem(name);
    assert.strictEqual(values, null);
  });
  it('Add named item', async () => {
    const namedItem = { name: 'bob', value: '$B2', comment: 'none' };
    await book.addNamedItem(namedItem.name, namedItem.value, namedItem.comment);
    assert.deepStrictEqual(sampleBook.namedItems[sampleBook.namedItems.length - 1], namedItem);
  });
  it('Add named item that already exists', async () => {
    const item = { name: 'alice', value: '$B2', comment: 'none' };
    await assert.rejects(async () => book.addNamedItem(item.name, item.value, item.comment), /Named item already exists/);
  });
  it('Delete named item', async () => {
    const name = 'alice';
    await book.deleteNamedItem(name);
    const index = sampleBook.namedItems.findIndex((item) => item.name === name);
    assert.strictEqual(index, -1);
  });
  it('Delete named item that doesn\'t exist', async () => {
    const name = 'fred';
    await assert.rejects(async () => book.deleteNamedItem(name), /Named item not found/);
  });
});
