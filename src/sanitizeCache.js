/*
 * Copyright 2026 Adobe. All rights reserved.
 * This file is licensed to you under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License. You may obtain a copy
 * of the License at http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under
 * the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
 * OF ANY KIND, either express or implied. See the License for the specific language
 * governing permissions and limitations under the License.
 */

// internals is not part of MSAL's public API — pin @azure/msal-node tightly
// and revisit if Deserializer is ever exposed publicly.
import { internals } from '@azure/msal-node';

/**
 * Derives the canonical MSAL cache key for a refresh token, matching the
 * format used by MSAL's internal serializer: segments are joined with `-` and
 * lowercased. Empty realm, target, and tokenType segments are included to keep
 * the key structure consistent with all other credential types.
 *
 * @param {object} refreshToken deserialized MSAL refresh-token entry
 * @param {string} refreshToken.homeAccountId account identifier
 * @param {string} refreshToken.environment authority host (e.g. `login.microsoftonline.com`)
 * @param {string} refreshToken.credentialType always `"RefreshToken"` for this entry type
 * @param {string} [refreshToken.familyId] family ID for first-party apps; takes precedence over
 *   clientId
 * @param {string} refreshToken.clientId application client ID
 * @returns {string} lowercase hyphen-joined cache key
 */
function generateKey(refreshToken) {
  const {
    homeAccountId, environment, credentialType, familyId, clientId,
  } = refreshToken;
  const credentialKey = [
    homeAccountId,
    environment,
    credentialType,
    familyId || clientId, // first-party apps share a family token; fall back to clientId
    '', // realm
    '', // target
    '', // tokenType
  ];
  return credentialKey.join('-').toLowerCase();
}

/**
 * Removes refresh-token entries whose cache key no longer matches the key that
 * MSAL would generate for them today. This repairs caches that were written by
 * an older MSAL version using a different key scheme.
 *
 * Mutates {@link data}.RefreshToken in place.
 *
 * If all tokens are outdated the cache is left untouched and a warning is
 * logged, because wiping every token would sign out all active sessions.
 *
 * @param {object} data raw MSAL serialized cache (JSON-parsed)
 * @param {{ info: Function, warn: Function }} log logger instance
 */
export function sanitizeCache(data, log) {
  const deserialized = internals.Deserializer.deserializeAllCache(data);
  const { refreshTokens } = deserialized;

  const outdatedKeys = [];
  for (const [actualKey, refreshToken] of Object.entries(refreshTokens)) {
    const expectedKey = generateKey(refreshToken);
    if (actualKey !== expectedKey) {
      outdatedKeys.push(actualKey);
    }
  }
  if (outdatedKeys.length === 0) {
    return;
  }

  // If every token is outdated the key scheme itself probably changed; removing
  // them all would revoke all sessions, so bail out and let the caller decide.
  if (Object.keys(refreshTokens).length === outdatedKeys.length) {
    log.warn('All refresh tokens have an unexpected key, generation might have changed.');
    return;
  }

  for (const key of outdatedKeys) {
    delete data.RefreshToken[key];
    log.info(`Removed RefreshToken with key: ${key}`);
  }
}
