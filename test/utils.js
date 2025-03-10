/*
 * Copyright 2022 Adobe. All rights reserved.
 * This file is licensed to you under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License. You may obtain a copy
 * of the License at http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under
 * the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
 * OF ANY KIND, either express or implied. See the License for the specific language
 * governing permissions and limitations under the License.
 */
import assert from 'assert';
import nock from 'nock';

export function Nock() {
  const DEFAULT_AUTH = {
    token_type: 'Bearer',
    refresh_token: 'dummy',
    access_token: 'dummy',
    expires_in: 181000,
  };

  const scopes = {};

  let unmatched;

  function noMatchHandler(req) {
    unmatched.push(req);
  }

  function nocker(url) {
    let scope = scopes[url];
    if (!scope) {
      scope = nock(url);
      scopes[url] = scope;
    }
    if (!unmatched) {
      unmatched = [];
      nock.emitter.on('no match', noMatchHandler);
    }
    return scope;
  }

  nocker.done = () => {
    Object.values(scopes).forEach((s) => s.done());
    if (unmatched) {
      assert.deepStrictEqual(unmatched.map((req) => req.options || req), []);
      nock.emitter.off('no match', noMatchHandler);
    }
  };

  nocker.loginWindowsNet = (auth = DEFAULT_AUTH) => nocker('https://login.windows.net')
    .post('/common/oauth2/token?api-version=1.0')
    .reply(200, auth);

  nocker.loginMicrosoftOnline = () => nocker('https://login.microsoftonline.com')
    .post('/adobe/oauth2/v2.0/token')
    .query((query) => {
      /* we only accept client-request-id or no query */
      if (query) {
        return Object.keys(query).every((key) => key === 'client-request-id');
      }
      return true;
    });

  nocker.token = (token) => nocker.loginMicrosoftOnline()
    .reply(200, token);

  nocker.unauthenticated = () => nocker.loginMicrosoftOnline()
    .reply(401, {
      error: 'invalid_client',
      error_description: 'AADSTS7000215: Invalid client secret provided.',
      error_codes: [
        7000215,
      ],
      timestamp: '2022-11-15 14:21:12Z',
      trace_id: '0360e583-c633-4ec7-a26d-691caf445c00',
      correlation_id: 'a498e2d2-2c57-41b3-a833-e361099aa522',
      error_uri: 'https://login.microsoftonline.com/error?code=7000215',
    });

  nocker.revoked = () => nocker.loginMicrosoftOnline()
    .reply(400, {
      error: 'invalid_grant',
      error_description: 'AADSTS50173: The provided grant has expired due to it being revoked, a fresh auth token is needed. The user might have changed or reset their password.',
      error_codes: [
        50173,
      ],
      timestamp: '2022-11-15 14:21:12Z',
      trace_id: '0360e583-c633-4ec7-a26d-691caf445c00',
      correlation_id: 'a498e2d2-2c57-41b3-a833-e361099aa522',
      error_uri: 'https://login.microsoftonline.com/error?code=50173',
      suberror: 'badtoken',
    });

  return nocker;
}
