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
import { UnsecuredJWT } from 'jose';
import { OneDrive } from '../src/OneDrive.js';
import { OneDriveAuth } from '../src/OneDriveAuth.js';
import { StatusCodeError } from '../src/StatusCodeError.js';
import { Nock } from './utils.js';
import { NamedItemContainer } from '../src/excel/NamedItemContainer.js';

/**
 * @param {OneDriveAuthOptions} opts
 * @returns {OneDriveAuth}
 */
const DEFAULT_AUTH = (opts = {}) => new OneDriveAuth({
  clientId: 'foo',
  localAuthCache: true,
  noTenantCache: true,
  ...opts,
}).setAccessToken(new UnsecuredJWT({
  email: 'bob',
  tid: 'test-tenantid',
}).encode());

describe('NamedItemContainer Tests', () => {
  let nock;
  beforeEach(() => {
    nock = new Nock();
    delete process.env.HELIX_ONEDRIVE_NO_SHARE_LINK_CACHE;
  });

  afterEach(() => {
    nock.done();
  });

  it('Get named item that returns another error than 404', async () => {
    nock('https://graph.microsoft.com/v1.0')
      .get('/names/test')
      .reply(500, 'Internal Server Error');

    const drive = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    const container = new NamedItemContainer(drive);
    container.uri = '';

    await assert.rejects(
      async () => container.getNamedItem('test'),
      new StatusCodeError('Internal Server Error', 500),
    );
  });

  it('Add named item that returns item exists but not as 409', async () => {
    const namedItem = { name: 'bob', value: '$B2', comment: 'none' };
    nock('https://graph.microsoft.com/v1.0')
      .post('/names/add')
      .reply(400, {
        message: 'Item already exists',
        code: 'ItemAlreadyExists',
      });

    const drive = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    const container = new NamedItemContainer(drive);
    container.uri = '';

    await assert.rejects(
      async () => container.addNamedItem(namedItem.name, namedItem.value, namedItem.comment),
      new StatusCodeError('Item already exists', 409),
    );
  });

  it('Delete named item that returns not found but not as 404', async () => {
    nock('https://graph.microsoft.com/v1.0')
      .delete('/names/test')
      .reply(400, {
        message: 'Item not found',
        code: 'ItemNotFound',
      });

    const drive = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    const container = new NamedItemContainer(drive);
    container.uri = '';

    await assert.rejects(
      async () => container.deleteNamedItem('test'),
      new StatusCodeError('Item not found', 404),
    );
  });
});
