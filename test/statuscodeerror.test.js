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

'use strict';

const assert = require('assert');
const StatusCodeError = require('../src/StatusCodeError.js');

describe('StatusCodeError Tests', () => {
  it('Set the status code', async () => {
    const error = new StatusCodeError('not found', 404);
    assert.equal(error.statusCode, 404);
  });

  it('fromError defaults to 500', async () => {
    const error = new Error('what the heck?!?');
    const e = StatusCodeError.fromError(error);
    assert.equal(e.statusCode, 500);
  });

  it('fromError gets details from inner error', async () => {
    const error = new Error('what the heck?!?');
    const e = StatusCodeError.fromErrorResponse(error, 404);
    assert.deepEqual(e.details, error);
  });

  it('fromErrorResponse gets details from inner error', async () => {
    const error = {
      code: 'itemNotFound',
    };
    const e = StatusCodeError.fromErrorResponse(error, 404);
    assert.equal(e.statusCode, 404);
    assert.deepEqual(e.details, { code: 'itemNotFound' });
  });
});
