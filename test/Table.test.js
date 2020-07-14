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

describe('Table Tests', () => {
  let table;
  let sampleTable;
  beforeEach(() => {
    const oneDrive = new MockOneDrive()
      .registerWorkbook('my-drive', 'my-item', exampleBook);
    const book = oneDrive.getWorkbook();
    table = book.table('table');
    [sampleTable] = oneDrive.workbooks[0].data.tables;
  });

  it('Rename a table', async () => {
    const name = 'table1';
    assert.equal(sampleTable.name, 'table');
    await table.rename(name);
    assert.equal(sampleTable.name, 'table1');
    console.log(exampleBook);
  });

  it('Get header names of table', async () => {
    const values = await table.getHeaderNames();
    assert.equal(values, sampleTable.headerNames);
  });

  it('Get all rows in table', async () => {
    const values = await table.getRows();
    assert.deepEqual(values, sampleTable.rows);
  });
  it('Get row in table', async () => {
    const index = 5;
    const values = await table.getRow(index);
    assert.deepEqual(values, sampleTable.rows[index]);
  });
  it('Get row in table with a bad index', async () => {
    const index = 20;
    await assert.rejects(async () => table.getRow(index), /Index out of range/);
  });
  it('Get rows as objects', async () => {
    const data = await table.getRowsAsObjects();
    assert.deepEqual(data, [
      { Firstname: 'Albert', Name: 'Einstein' },
      { Firstname: 'Marie', Name: 'Curie' },
      { Firstname: 'Steven', Name: 'Hawking' },
      { Firstname: 'Isaac', Name: 'Newton' },
      { Firstname: 'Niels', Name: 'Bohr' },
      { Firstname: 'Michael', Name: 'Faraday' },
      { Firstname: 'Galileo', Name: 'Galilei' },
      { Firstname: 'Johannes', Name: 'Kepler' },
      { Firstname: 'Nikolaus', Name: 'Kopernikus' },
    ]);
  });
  it('Add row to table', async () => {
    const row = ['Heisenberg', 'Werner'];
    const index = await table.addRow(row);
    assert.deepEqual(row, sampleTable.rows[index]);
  });
  it('Add rows to table', async () => {
    const rows = [['Heisenberg', 'Werner'], ['Planck', 'Max']];
    const index = await table.addRows(rows);
    assert.deepEqual(rows[0], sampleTable.rows[index - 1]);
    assert.deepEqual(rows[1], sampleTable.rows[index]);
  });
  it('Replace row in table', async () => {
    const index = 5;
    const row = ['Heisenberg', 'Werner'];
    await table.replaceRow(index, row);
    assert.deepEqual(row, sampleTable.rows[index]);
  });
  it('Get number of rows it table', async () => {
    const count = await table.getRowCount();
    assert.equal(count, sampleTable.rows.length);
  });
  it('Get column in table', async () => {
    const index = 0;
    const values = await table.getColumn('Name');
    assert.deepEqual(values[0], [sampleTable.headerNames[index]]);
    assert.deepEqual(values[5], [sampleTable.rows[4][index]]);
  });
});
