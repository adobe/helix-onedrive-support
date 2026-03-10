/*
 * Copyright 2019 Adobe. All rights reserved.
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
import { keepAliveNoCache } from '@adobe/fetch';
import { CustomNetworkModule } from '../src/CustomNetworkModule.js';
import { Nock } from './utils.js';

const AZ_AUTHORITY_HOST_URL = 'https://login.windows.net';

describe('CustomNetworkModule Tests', () => {
  let nock;
  let fetchContext;
  let networkModule;

  beforeEach(() => {
    nock = new Nock();
    fetchContext = keepAliveNoCache({ userAgent: 'adobe-fetch' });
    networkModule = new CustomNetworkModule(fetchContext);
  });

  afterEach(async () => {
    await fetchContext.reset();
    nock.done();
  });

  it('send a GET request', async () => {
    nock(AZ_AUTHORITY_HOST_URL)
      .get('/')
      .reply(200, { message: 'Hello, world!' });

    const response = await networkModule.sendGetRequestAsync(AZ_AUTHORITY_HOST_URL);
    assert.deepStrictEqual(response, {
      status: 200,
      headers: {
        'content-type': 'application/json',
      },
      body: {
        message: 'Hello, world!',
      },
    });
  });

  it('send a POST request', async () => {
    nock(AZ_AUTHORITY_HOST_URL)
      .post('/')
      .reply(200, { message: 'Hello, world!' });

    const response = await networkModule.sendPostRequestAsync(AZ_AUTHORITY_HOST_URL);
    assert.deepStrictEqual(response, {
      status: 200,
      headers: {
        'content-type': 'application/json',
      },
      body: {
        message: 'Hello, world!',
      },
    });
  });

  it('send a request that times out', async () => {
    nock(AZ_AUTHORITY_HOST_URL)
      .get('/')
      .delay(1000)
      .reply(200, { message: 'Hello, world!' });

    await assert.rejects(
      () => networkModule.sendGetRequestAsync(AZ_AUTHORITY_HOST_URL, {}, 10),
      /Request timeout/,
    );
  });

  it('send a request that throws an error', async () => {
    nock(AZ_AUTHORITY_HOST_URL)
      .get('/')
      .replyWithError(new Error('boom!'));

    await assert.rejects(
      () => networkModule.sendGetRequestAsync(AZ_AUTHORITY_HOST_URL),
      /network_error: Network request failed: boom!/,
    );
  });

  it('return a response that is not JSON', async () => {
    nock(AZ_AUTHORITY_HOST_URL)
      .get('/')
      .reply(200, 'not json');

    await assert.rejects(
      () => networkModule.sendGetRequestAsync(AZ_AUTHORITY_HOST_URL),
      /token_parsing_error: Failed to parse response: Unexpected token 'o'/,
    );
  });
});
