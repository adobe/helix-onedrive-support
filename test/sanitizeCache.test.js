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
import { internals } from '@azure/msal-node';
import { sanitizeCache } from '../src/sanitizeCache.js';

/**
 * Creates a minimal logger spy whose recorded calls are exposed via `.calls`.
 *
 * @returns {{ info: Function, warn: Function, calls: { info: string[], warn: string[] } }}
 */
function makeLog() {
  const calls = { info: [], warn: [] };
  return {
    info: (msg) => { calls.info.push(msg); },
    warn: (msg) => { calls.warn.push(msg); },
    calls,
  };
}

/**
 * Builds a serialized MSAL cache object containing only refresh tokens.
 * Accepts camelCase token entities (as used internally by MSAL) and delegates
 * serialization to {@link internals.Serializer.serializeRefreshTokens}.
 *
 * @param {Record<string, object>} tokens map of cache key to camelCase RefreshTokenEntity
 * @returns {{ RefreshToken: Record<string, object> }} serialized cache fragment
 */
function makeCache(tokens) {
  return {
    RefreshToken: internals.Serializer.serializeRefreshTokens(tokens),
  };
}

/**
 * Derives the expected MSAL cache key for a refresh token. Mirrors the
 * `generateKey` logic in the source so tests remain independent of it.
 *
 * @param {object} token
 * @param {string} token.homeAccountId
 * @param {string} token.environment
 * @param {string} token.credentialType
 * @param {string} token.clientId
 * @param {string} [token.familyId]
 * @returns {string}
 */
function canonicalKey({
  homeAccountId, environment, credentialType, clientId, familyId,
}) {
  return [homeAccountId, environment, credentialType, familyId || clientId, '', '', '']
    .join('-')
    .toLowerCase();
}

const BASE = {
  homeAccountId: 'uid.tenantid',
  environment: 'login.microsoftonline.com',
  credentialType: 'RefreshToken',
  clientId: 'my-client-id',
};

describe('sanitizeCache', () => {
  it('does nothing when there are no refresh tokens', () => {
    const data = {};
    const log = makeLog();
    sanitizeCache(data, log);
    assert.deepStrictEqual(log.calls.info, []);
    assert.deepStrictEqual(log.calls.warn, []);
  });

  it('does nothing when all keys are already correct', () => {
    const key = canonicalKey(BASE);
    const data = makeCache({ [key]: { ...BASE, secret: 'dummy' } });
    const log = makeLog();
    sanitizeCache(data, log);
    assert.deepStrictEqual(Object.keys(data.RefreshToken), [key]);
    assert.deepStrictEqual(log.calls.info, []);
    assert.deepStrictEqual(log.calls.warn, []);
  });

  it('leaves a single outdated token and warns when it is the only entry', () => {
    // A cache with one token whose key is stale counts as "all outdated" — the
    // bail-out guard fires rather than deleting the sole session.
    const staleKey = `stale-${canonicalKey(BASE)}`;
    const data = makeCache({ [staleKey]: { ...BASE, secret: 'dummy' } });
    const log = makeLog();
    sanitizeCache(data, log);
    assert.ok(staleKey in data.RefreshToken, 'sole token must not be removed');
    assert.strictEqual(log.calls.warn.length, 1);
    assert.deepStrictEqual(log.calls.info, []);
  });

  it('removes only outdated tokens when the cache has a mix of valid and stale entries', () => {
    const validKey = canonicalKey(BASE);
    const staleKey = 'old-format-key';
    const data = makeCache({
      [validKey]: { ...BASE, secret: 'dummy' },
      [staleKey]: { ...BASE, secret: 'dummy' },
    });
    const log = makeLog();
    sanitizeCache(data, log);
    assert.ok(validKey in data.RefreshToken, 'valid token must be kept');
    assert.strictEqual(data.RefreshToken[staleKey], undefined);
    assert.deepStrictEqual(log.calls.warn, []);
  });

  it('leaves all tokens untouched and emits a warning when every key is outdated', () => {
    const staleKey1 = 'stale-key-1';
    const staleKey2 = 'stale-key-2';
    const data = makeCache({
      [staleKey1]: { ...BASE, secret: 'dummy' },
      [staleKey2]: { ...BASE, clientId: 'other-client', secret: 'dummy' },
    });
    const log = makeLog();
    sanitizeCache(data, log);
    assert.ok(staleKey1 in data.RefreshToken);
    assert.ok(staleKey2 in data.RefreshToken);
    assert.strictEqual(log.calls.warn.length, 1);
    assert.deepStrictEqual(log.calls.info, []);
  });

  it('uses familyId instead of clientId when building the key for first-party tokens', () => {
    const firstParty = { ...BASE, familyId: '1' };
    const key = canonicalKey(firstParty);
    const data = makeCache({ [key]: { ...firstParty, secret: 'dummy' } });
    const log = makeLog();
    sanitizeCache(data, log);
    assert.deepStrictEqual(Object.keys(data.RefreshToken), [key]);
    assert.deepStrictEqual(log.calls.info, []);
  });

  it('logs one info message per removed token', () => {
    const validKey = canonicalKey(BASE);
    const stale1 = 'stale-1';
    const stale2 = 'stale-2';
    const data = makeCache({
      [validKey]: { ...BASE, secret: 'dummy' },
      [stale1]: { ...BASE, secret: 'dummy' },
      [stale2]: { ...BASE, secret: 'dummy' },
    });
    const log = makeLog();
    sanitizeCache(data, log);
    assert.strictEqual(log.calls.info.length, 2);
    assert.ok(log.calls.info.some((m) => m.includes(stale1)));
    assert.ok(log.calls.info.some((m) => m.includes(stale2)));
  });
});
