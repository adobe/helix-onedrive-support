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
import { S3CachePlugin } from '@adobe/helix-shared-tokencache';

async function run() {
  const contentBusId = process.argv[2];
  if (!contentBusId) {
    console.error('usage clear-access-token <contentBusId> [expire]');
    process.exit(1);
  }
  const p = new S3CachePlugin({
    bucket: 'helix-content-bus',
    key: `${contentBusId}/.helix-auth/auth-onedrive-content.json`,
    secret: contentBusId,
  });
  let data = {};
  const ctx = {
    tokenCache: {
      deserialize(json) {
        data = JSON.parse(json);
      },
      serialize() {
        return JSON.stringify(data);
      },
    },
  };
  await p.beforeCacheAccess(ctx);
  if (process.argv[3]) {
    const at = data.AccessToken[Object.keys(data.AccessToken)[0]];
    at.cached_at = 1;
    at.expires_on = 2;
    at.extended_expires_on = 3;
  } else {
    delete data.AccessToken;
  }
  ctx.cacheHasChanged = true;
  await p.afterCacheAccess(ctx);
  console.log(`Account: ${Object.values(data.Account)[0].username}`);
  if (process.argv[3]) {
    console.log('expired access token');
  } else {
    console.log('cleared access token');
  }
}

run().catch(console.error);
