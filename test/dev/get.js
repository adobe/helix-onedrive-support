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

/* eslint-disable no-console */

import { config } from 'dotenv';
import { S3CacheManager } from '@adobe/helix-shared-tokencache';
import { GetObjectCommand, S3Client } from '@aws-sdk/client-s3';
import { promisify } from 'util';
import zlib from 'zlib';
import { Response } from '@adobe/fetch';

const gunzip = promisify(zlib.gunzip);

config();

function createCacheContext() {
  return {
    tokenCache: {
      deserialize(json) {
        const data = JSON.parse(json);
        console.log(data);
        const accounts = Object.values(data.Account);
        if (accounts.length > 0) {
          console.log(`\n\nAccount: ${accounts[0].username}`);
        } else {
          console.log('\n\nNo Accounts.');
        }
        if (data.AccessToken) {
          const accessToken = Object.values(data.AccessToken)[0];
          console.log(`Access token expires on: ${new Date(Number(accessToken.expires_on) * 1000).toISOString()}`);
          console.log(`Access token cached at:  ${new Date(Number(accessToken.cached_at) * 1000).toISOString()}`);
        } else {
          console.log('no access token');
        }
      },

    },
  };
}

async function run() {
  const [, , owner, repo, type = 'onedrive'] = process.argv;
  if (!owner) {
    console.error('usage: node get owner repo [type = onedrive]');
    process.exit(1);
  }

  let contentBusId;
  if (owner === 'default') {
    contentBusId = 'default';
  } else {
    try {
      const s3 = new S3Client();
      const res = await s3.send(new GetObjectCommand({
        Bucket: 'helix-code-bus',
        Key: `${owner}/${repo}/main/helix-config.json`,
      }));
      contentBusId = res.Metadata['x-contentbus-id'].substring(2);
    } catch (e) {
      console.error(`unable to load helix-config.json:${e.message}`);
    }
  }
  if (!contentBusId) {
    // load from helix5 config
    try {
      const s3 = new S3Client();
      const res = await s3.send(new GetObjectCommand({
        Bucket: 'helix-config-bus',
        Key: `orgs/${owner}/sites/${repo}.json`,
      }));
      let buf = await new Response(res.Body, {}).buffer();
      if (res.ContentEncoding === 'gzip') {
        buf = await gunzip(buf);
      }
      contentBusId = JSON.parse(buf).content.contentBusId;
    } catch (e) {
      console.error(`unable to load helix5 config:${e.message}`);
    }
  }
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

  console.log('project cache');
  console.log('-----------------------------------');
  if (await projectCache.hasCache('content')) {
    const p = await projectCache.getCache('content');
    await p.beforeCacheAccess(createCacheContext());
    console.log('plugin metadata:');
    console.log(await p.getPluginMetadata());
  } else {
    console.log('n/a');
  }

  console.log('\norg cache');
  console.log('-----------------------------------');
  if (await orgCache.hasCache('content')) {
    const p = await orgCache.getCache('content');
    await p.beforeCacheAccess(createCacheContext());
  } else {
    console.log('n/a');
  }
}

run().catch(console.error);
