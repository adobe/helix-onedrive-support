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

describe('Table Tests', () => {
  /** @type {import('../src/index.js').Table} */
  let table;
  let sampleTable;
  let book;
  beforeEach(() => {
    const oneDrive = new OneDriveMock()
      .registerWorkbook('my-drive', 'my-item', exampleBook);
    book = oneDrive.getWorkbook();
    table = book.table('table');
    [sampleTable] = oneDrive.workbooks[0].data.tables;
  });

  it('Rename a table', async () => {
    const name = 'table1';
    assert.strictEqual(sampleTable.name, 'table');
    await table.rename(name);
    assert.strictEqual(sampleTable.name, 'table1');
  });

  it('Rename a table that does not exist', async () => {
    const name = 'table1';
    table = book.table('table-not-exist');
    await assert.rejects(async () => table.rename(name), new StatusCodeError('table-not-exist', 404));
  });

  it('Get header names of table', async () => {
    const values = await table.getHeaderNames();
    assert.strictEqual(values, sampleTable.headerNames);
  });

  it('Get all rows in table', async () => {
    const values = await table.getRows();
    assert.deepStrictEqual(values, sampleTable.rows);
  });
  it('Get row in table', async () => {
    const index = 5;
    const values = await table.getRow(index);
    assert.deepStrictEqual(values, sampleTable.rows[index]);
  });
  it('Get row in table with a bad index', async () => {
    const index = 20;
    await assert.rejects(async () => table.getRow(index), /Index out of range/);
  });
  it('Get rows as objects', async () => {
    const data = await table.getRowsAsObjects();
    assert.deepStrictEqual(data, [
      { ' F i r s t n a m e ': 'Albert', Name: 'Einstein' },
      { ' F i r s t n a m e ': 'Marie', Name: 'Curie' },
      { ' F i r s t n a m e ': 'Steven', Name: 'Hawking' },
      { ' F i r s t n a m e ': 'Isaac', Name: 'Newton' },
      { ' F i r s t n a m e ': 'Niels', Name: 'Bohr' },
      { ' F i r s t n a m e ': 'Michael', Name: 'Faraday' },
      { ' F i r s t n a m e ': 'Galileo', Name: 'Galilei' },
      { ' F i r s t n a m e ': 'Johannes', Name: 'Kepler' },
      { ' F i r s t n a m e ': 'Nikolaus', Name: 'Kopernikus' },
      { ' F i r s t n a m e ': '\t Balls ', Name: ' Space\u200B' },
    ]);
  });
  it('Get rows as objects (trimmed)', async () => {
    const data = await table.getRowsAsObjects({ trim: true });
    assert.deepStrictEqual(data, [
      { 'F i r s t n a m e': 'Albert', Name: 'Einstein' },
      { 'F i r s t n a m e': 'Marie', Name: 'Curie' },
      { 'F i r s t n a m e': 'Steven', Name: 'Hawking' },
      { 'F i r s t n a m e': 'Isaac', Name: 'Newton' },
      { 'F i r s t n a m e': 'Niels', Name: 'Bohr' },
      { 'F i r s t n a m e': 'Michael', Name: 'Faraday' },
      { 'F i r s t n a m e': 'Galileo', Name: 'Galilei' },
      { 'F i r s t n a m e': 'Johannes', Name: 'Kepler' },
      { 'F i r s t n a m e': 'Nikolaus', Name: 'Kopernikus' },
      { 'F i r s t n a m e': 'Balls', Name: 'Space' },
    ]);
  });
  it('Add row to table', async () => {
    const row = ['Heisenberg', 'Werner'];
    const index = await table.addRow(row);
    assert.deepStrictEqual(row, sampleTable.rows[index]);
  });
  it('Add rows to table', async () => {
    const rows = [['Heisenberg', 'Werner'], ['Planck', 'Max']];
    const index = await table.addRows(rows);
    assert.deepStrictEqual(rows[0], sampleTable.rows[index - 1]);
    assert.deepStrictEqual(rows[1], sampleTable.rows[index]);
  });
  it('Replace row in table', async () => {
    const index = 5;
    const row = ['Heisenberg', 'Werner'];
    await table.replaceRow(index, row);
    assert.deepStrictEqual(row, sampleTable.rows[index]);
  });
  it('Get number of rows it table', async () => {
    const count = await table.getRowCount();
    assert.strictEqual(count, sampleTable.rows.length);
  });
  it('Get column in table', async () => {
    const index = 0;
    const values = await table.getColumn('Name');
    assert.deepStrictEqual(values[0], [sampleTable.headerNames[index]]);
    assert.deepStrictEqual(values[5], [sampleTable.rows[4][index]]);
  });
  it('Get column in table that does not exist', async () => {
    await assert.rejects(table.getColumn('Foobar'), new StatusCodeError('Column name not found: Foobar', 400));
  });
  it('Add column to table', async () => {
    await table.addColumn('newHeader');
    const headerNames = await table.getHeaderNames();
    assert.strictEqual(headerNames.indexOf('newHeader'), headerNames.length - 1);
  });
  it('Add column to table at front', async () => {
    await table.addColumn('newHeader', 0);
    const headerNames = await table.getHeaderNames();
    assert.strictEqual(headerNames.indexOf('newHeader'), 0);
  });
  it('Delete column from table', async () => {
    await table.deleteColumn('Name');
    const headerNames = await table.getHeaderNames();
    assert.strictEqual(headerNames.indexOf('Name'), -1);
  });
  it('Delete column from table that does not exist', async () => {
    await assert.rejects(table.deleteColumn('Foobar'), new StatusCodeError('Column name not found: Foobar', 400));
  });
  it('Delete row in table', async () => {
    const index = 5;
    const rowAfter = sampleTable.rows[index + 1];
    await table.deleteRow(index);
    assert.deepStrictEqual(rowAfter, sampleTable.rows[index]);
  });
});
