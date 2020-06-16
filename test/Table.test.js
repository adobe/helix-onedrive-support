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

const StatusCodeError = require('../src/StatusCodeError.js');
const Table = require('../src/Table.js');

const getClient = require('./getClient.js');

const sampleTable = {
  name: 'table',
  headerNames: ['Name', 'Firstname'],
  rows: [
    ['Einstein', 'Albert'],
    ['Curie', 'Marie'],
    ['Hawking', 'Steven'],
    ['Newton', 'Isaac'],
    ['Bohr', 'Niels'],
    ['Faraday', 'Michael'],
    ['Galilei', 'Galileo'],
    ['Kepler', 'Johannes'],
    ['Kopernikus', 'Nikolaus'],
  ],
  ops: ({
    component, name, command, body,
  }) => {
    let index;
    switch (component) {
      case 'dataBodyRange':
        return { rowCount: sampleTable.rows.length };
      case 'headerRowRange':
        return { values: [sampleTable.headerNames] };
      case 'rows':
        if (!command) {
          return { value: sampleTable.rows.map((row) => ({ values: [row] })) };
        }
        if (command === 'add') {
          sampleTable.rows.push(body.values[0]);
          return { index: sampleTable.rows.length - 1 };
        }
        index = parseInt(command.replace(/itemAt\(index=([0-9]+)\)/, '$1'), 10);
        if (index < 0 || index >= sampleTable.rows.length) {
          throw new StatusCodeError(`Index out of range: ${index}`, 400);
        }
        if (body) {
          [sampleTable.rows[index]] = body.values;
        }
        return { values: [sampleTable.rows[index]] };
      case 'columns':
        index = sampleTable.headerNames.findIndex((n) => n === name);
        if (index === -1) {
          throw new StatusCodeError(`Column name not found: ${name}`, 400);
        }
        return {
          values: [
            [sampleTable.headerNames[index]],
            ...sampleTable.rows.map((row) => [row[index]]),
          ],
        };
      default:
        if (body) {
          sampleTable.name = body.name;
        }
        return { values: sampleTable.name };
    }
  },
};

const oneDrive = {
  getClient: async () => getClient(sampleTable.ops),
};

describe('Table Tests', () => {
  const table = new Table(oneDrive, 'workbook', 'table', console);
  it('Rename a table', async () => {
    const name = 'table1';
    await table.rename(name);
    assert.equal(name, sampleTable.name);
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
  it('Add row to table', async () => {
    const row = ['Heisenberg', 'Werner'];
    const index = await table.addRow(row);
    assert.deepEqual(row, sampleTable.rows[index]);
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
