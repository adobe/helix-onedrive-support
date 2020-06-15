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
const Worksheet = require('../src/Worksheet.js');

const getClient = require('./getClient.js');
const namedItemOps = require('./NamedItemOps.js');

const sampleSheet = {
  name: 'sheet',
  namedItems: [
    { name: 'alice', value: '$A2', comment: 'none' },
  ],
  ops: ({
    component, command, method, body,
  }) => {
    switch (component) {
      case 'names':
        return namedItemOps(sampleSheet.namedItems)({ command, method, body });
      default:
        return { values: sampleSheet.name };
    }
  },
};

const oneDrive = {
  getClient: async () => getClient(sampleSheet.ops),
};

describe('Worksheet Tests', () => {
  const sheet = new Worksheet(oneDrive, 'workbook', 'sheet', console);
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
});
