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
  const contentBusId = process.argv[2] || 'default';
  const p = new S3CachePlugin({
    bucket: 'helix-content-bus',
    key: `${contentBusId}/.helix-auth/auth-onedrive-content.json`,
    secret: contentBusId,
  });
  const ctx = {
    tokenCache: {
      deserialize(json) {
        const data = JSON.parse(json);
        console.log(data);
        console.log(`\n\nAccount: ${Object.values(data.Account)[0].username}`);
        if (data.AccessToken) {
          const accessToken = Object.values(data.AccessToken)[0];
          console.log(`Access token expires on: ${new Date(Number(accessToken.expires_on) * 1000).toISOString()}`);
        } else {
          console.log('no access token');
        }
      },
    },
  };
  await p.beforeCacheAccess(ctx);
}

run().catch(console.error);
