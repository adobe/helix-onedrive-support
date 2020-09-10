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

  it('getActualError unwraps the underlying error', async () => {
    const origError = new StatusCodeError('not found', 404);
    const newError = new Error('An error happened');
    newError.error = origError;
    assert.equal(StatusCodeError.getActualError(newError), origError);
  });

  it('getActualError is robust against `undefined`', async () => {
    const error = new StatusCodeError('not found', 404);
    error.error = undefined;
    assert.equal(StatusCodeError.getActualError(error), error);
  });

  it('getActualError is robust against `null`', async () => {
    const error = new StatusCodeError('not found', 404);
    error.error = null;
    assert.equal(StatusCodeError.getActualError(error), error);
  });
});
