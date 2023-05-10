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
import { S3CacheManager } from '@adobe/helix-shared-tokencache';
import { GetObjectCommand, S3Client } from '@aws-sdk/client-s3';

async function run() {
  const [, , owner, repo, type = 'onedrive', key = 'content'] = process.argv;
  if (!owner) {
    console.error('usage: node get owner repo [type = onedrive] [key = content]');
    process.exit(1);
  }
  const s3 = new S3Client();
  const res = await s3.send(new GetObjectCommand({
    Bucket: 'helix-code-bus',
    Key: `${owner}/${repo}/main/helix-config.json`,
  }));
  const contentBusId = res.Metadata['x-contentbus-id'].substring(2);
  if (!contentBusId) {
    throw Error('no contentBusId');
  }

  const projectCache = new S3CacheManager({
    log: console,
    prefix: `${contentBusId}/.helix-auth`,
    secret: contentBusId,
    bucket: 'helix-content-bus',
    type,
  });
  const orgCache = new S3CacheManager({
    log: console,
    prefix: `${owner}/.helix-auth`,
    secret: owner,
    bucket: 'helix-code-bus',
    type,
  });

  if (!await projectCache.hasCache(key)) {
    console.error('project has no tokencache');
    process.exit(1);
  }

  if (await orgCache.hasCache(key)) {
    console.warn(('overwriting existing org cache!'));
  }

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

  const projectPlugin = await projectCache.getCache(key);
  const orgPlugin = await orgCache.getCache(key);

  await projectPlugin.beforeCacheAccess(ctx);
  ctx.cacheHasChanged = true;
  await orgPlugin.afterCacheAccess(ctx);
  console.log(`Account: ${Object.values(data.Account)[0].username}`);
  console.log('token updated.');
}

run().catch(console.error);
