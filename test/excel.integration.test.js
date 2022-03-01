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
const assert = require('assert');

const OneDrive = require('../src/OneDrive.js');

require('dotenv').config();

describe('Excel Integration Tests', () => {
  it('Get the sheet data', async function test() {
    if (!process.env.AZURE_WORD2MD_CLIENT_ID) {
      this.skip();
      return;
    }
    const drive = new OneDrive({
      clientId: process.env.AZURE_WORD2MD_CLIENT_ID,
      username: process.env.AZURE_HELIX_USER,
      password: process.env.AZURE_HELIX_PASSWORD,
    });

    const rootItem = await drive.getDriveItemFromShareLink('https://adobe.sharepoint.com/sites/cg-helix/Shared%20Documents/helix-test-content-onedrive/automation-tests');
    const items = await drive.fuzzyGetDriveItem(rootItem, encodeURI('/pet-shop'));
    const book = await drive.getWorkbook(items[0]);
    const names = await book.getWorksheetNames();
    assert.deepStrictEqual(names, ['helix-default', 'incoming', 'Config']);
  }).timeout(10000);

  it('Test pre authenticate fetch', async function test() {
    if (!process.env.AZURE_WORD2MD_CLIENT_ID) {
      this.skip();
      return;
    }
    const drive = new OneDrive({
      clientId: process.env.AZURE_WORD2MD_CLIENT_ID,
      username: process.env.AZURE_HELIX_USER,
      password: process.env.AZURE_HELIX_PASSWORD,
    });

    await drive.getAccessToken();
    const books = ['/pet-shop', '/load-test', '/doccloud-test'];
    const result = await Promise.all(books.map(async (bookName) => {
      const rootItem = await drive.getDriveItemFromShareLink('https://adobe.sharepoint.com/sites/cg-helix/Shared%20Documents/helix-test-content-onedrive/automation-tests');
      const items = await drive.fuzzyGetDriveItem(rootItem, encodeURI(bookName));
      const book = await drive.getWorkbook(items[0]);
      // eslint-disable-next-line no-return-await
      return await book.getWorksheetNames();
    }));

    assert.deepStrictEqual(result, [
      ['helix-default', 'incoming', 'Config'],
      ['Sheet1', 'Config'],
      ['Sheet1', 'Config'],
    ]);
  }).timeout(10000);
});
