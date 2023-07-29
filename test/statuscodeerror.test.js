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
import { AbortError, FetchError } from '@adobe/fetch';
import { StatusCodeError } from '../src/index.js';

describe('StatusCodeError Tests', () => {
  it('Set the status code', async () => {
    const error = new StatusCodeError('not found', 404);
    assert.strictEqual(error.statusCode, 404);
  });

  it('fromError defaults to 500', async () => {
    const error = new Error('what the heck?!?');
    const e = StatusCodeError.fromError(error);
    assert.strictEqual(e.statusCode, 500);
  });

  it('fromError gets details from inner error', async () => {
    const error = new Error('what the heck?!?');
    const e = StatusCodeError.fromErrorResponse(error, 404);
    assert.deepStrictEqual(e.details, error);
  });

  it('fromErrorResponse gets details from inner error', async () => {
    const error = {
      code: 'itemNotFound',
    };
    const e = StatusCodeError.fromErrorResponse(error, 404);
    assert.strictEqual(e.statusCode, 404);
    assert.deepStrictEqual(e.details, { code: 'itemNotFound' });
  });

  it('fromErrorResponse gets message from inner error', async () => {
    const error = {
      message: 'iten was not found',
    };
    const e = StatusCodeError.fromErrorResponse({ error }, 404);
    assert.strictEqual(e.statusCode, 404);
    assert.deepStrictEqual(e.details, { message: error.message });
  });

  it('fromError recognizes AbortError', async () => {
    const error = new AbortError('aborted');
    const e = StatusCodeError.fromError(error);
    assert.strictEqual(e.statusCode, 504);
  });

  it('fromError recognizes FetchError', async () => {
    let error = new FetchError('whoops!', 'system', { code: 'ECONNRESET' });
    let e = StatusCodeError.fromError(error);
    assert.strictEqual(e.statusCode, 504);

    error = new FetchError('whoops!', 'system', { code: 'ETIMEDOUT' });
    e = StatusCodeError.fromError(error);
    assert.strictEqual(e.statusCode, 504);
  });
});
