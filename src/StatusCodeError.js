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

import { AbortError, FetchError } from '@adobe/fetch';

/**
 * Internal error class
 * @private
 */
export class StatusCodeError extends Error {
  /**
   * Converts a fetch error to a status code error.
   * @param {Error} err The original error
   * @returns {StatusCodeError} status code error
   */
  static fromError(err) {
    let statusCode = 500;
    if (err instanceof AbortError) {
      statusCode = 504;
    }
    if (err instanceof FetchError) {
      if (err.code === 'ECONNRESET' || err.code === 'ETIMEDOUT') {
        statusCode = 504;
      }
    }
    return new StatusCodeError(err.message, statusCode, err);
  }

  /**
   * Converts a Graph API error response to a status code error.
   * @param {object} errorBody The parsed error response body
   * @param {number} statusCode The status code of the error response
   * @param {RateLimit} rateLimit rate limit or null
   * @param {object} details The underlying error
   * @returns {StatusCodeError} status code error
   */
  static fromErrorResponse(errorBody, statusCode, rateLimit) {
    if (errorBody.error && errorBody.error.message) {
      // eslint-disable-next-line no-param-reassign
      errorBody = errorBody.error;
    }
    return new StatusCodeError(errorBody.message, statusCode, errorBody, rateLimit);
  }

  /**
   * Constructs a ne StatusCodeError.
   * @constructor
   * @param {string} msg Error message
   * @param {number} statusCode Status code of the error response
   * @param {object} details underlying error
   * @param {RateLimit} rateLimit rate limit or null
   */
  constructor(msg, statusCode, details, rateLimit) {
    super(msg?.value ?? msg);
    this.statusCode = statusCode;
    this.details = details;
    this.rateLimit = rateLimit;
  }
}
