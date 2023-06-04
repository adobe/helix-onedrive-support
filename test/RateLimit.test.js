/*
 * Copyright 2023 Adobe. All rights reserved.
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
import { Headers } from '@adobe/fetch';
import { Nock } from './utils.js';
import { RateLimit } from '../src/RateLimit.js';

describe('RateLimit Tests', () => {
  let nock;
  beforeEach(() => {
    nock = new Nock();
    delete process.env.HELIX_ONEDRIVE_NO_SHARE_LINK_CACHE;
  });

  afterEach(() => {
    nock.done();
  });

  it('Creates no RateLimit object when no headers are present', async () => {
    const headers = new Headers();
    const rateLimit = RateLimit.fromHeaders(headers);
    assert.strictEqual(rateLimit, null);
  });

  it('Creates RateLimit object when some headers are present', async () => {
    const headers = new Headers();
    headers.set('RateLimit-Limit', 10);
    const rateLimit = RateLimit.fromHeaders(headers);
    assert.notStrictEqual(rateLimit?.toString(), null);
  });

  it('Returns retryAfter when header is present', async () => {
    const headers = new Headers();
    headers.set('Retry-After', 30);
    const rateLimit = RateLimit.fromHeaders(headers);
    assert.notStrictEqual(rateLimit?.toString(), null);
  });
});
