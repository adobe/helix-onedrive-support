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

/**
 * Rate limit class
 */
export class RateLimit {
  static fromHeaders(headers) {
    const names = [
      ['RateLimit-Limit', 'limit'],
      ['RateLimit-Remaining', 'remaining'],
      ['RateLimit-Reset', 'reset'],
      ['Retry-After', 'retryAfter'],
    ];

    const result = {};
    let notEmpty = false;

    names.forEach(([hdr, prop]) => {
      const valueS = headers.get(hdr);
      if (valueS) {
        const value = Number.parseInt(valueS, 10);
        if (!Number.isNaN(value)) {
          result[prop] = value;
          notEmpty = true;
        }
      }
    });
    return notEmpty ? new RateLimit(result) : null;
  }

  constructor({
    limit, remaining, reset, retryAfter,
  }) {
    this._limit = limit;
    this._remaining = remaining;
    this._reset = reset;
    this._retryAfter = retryAfter;
  }

  get limit() {
    return this._limit;
  }

  get remaining() {
    return this._remaining;
  }

  get reset() {
    return this._reset;
  }

  get retryAfter() {
    return this._retryAfter;
  }

  toJSON() {
    const o = {
      limit: this.limit,
      remaining: this.remaining,
      reset: this.reset,
    };
    if (this.retryAfter) {
      o.retryAfter = this.retryAfter;
    }
    return o;
  }
}
