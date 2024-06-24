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
import { MemCachePlugin, S3CachePlugin } from '@adobe/helix-shared-tokencache';
import chalk from 'chalk-template';
import 'dotenv/config';
import readlinePromises from 'readline/promises';
import { OneDrive, OneDriveAuth } from '../../src/index.js';

const rl = readlinePromises.createInterface({
  input: process.stdin,
  output: process.stdout,
});

async function testReadAccessOnedrive(client, url) {
  try {
    const root = await client.resolveShareLink(url);
    console.log(chalk`\n{yellow access validated. user can access }{blue ${root.webUrl}}\n`);
  } catch (e) {
    console.warn(chalk`{red unable to resolve sharelink}`, e);
    if (e.details.code === 'accessDenied') {
      return `The resource specified in the fstab.yaml does either not exist, or you do not have permission to access it. Please make sure that the url is correct, the enterprise application: "Franklin Registration Service (${client.auth.clientId})" is consented for the required scopes, and that the logged in user has permissions to access it.`;
    } else {
      return `Unable to validate access: ${e.message}`;
    }
  }
  return '';
}

async function clearAccessToken(plugin, expire) {
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
  await plugin.beforeCacheAccess(ctx);
  if (expire) {
    const at = data.AccessToken?.[Object.keys(data.AccessToken)[0]];
    if (!at) {
      console.log(chalk`{yellow no access tokens}`);
      return;
    }
    at.cached_at = 1;
    at.expires_on = 2;
    at.extended_expires_on = 3;
  } else {
    delete data.AccessToken;
  }
  ctx.cacheHasChanged = true;
  await plugin.afterCacheAccess(ctx);
  if (expire) {
    console.log(chalk`{green expired access token}`);
  } else {
    console.log(chalk`{green cleared access token}`);
  }
}

async function getClient(cachePlugin, url) {
  const auth = new OneDriveAuth({
    log: console,
    clientId: process.env.AZURE_HELIX_SERVICE_CLIENT_ID,
    clientSecret: process.env.AZURE_HELIX_SERVICE_CLIENT_SECRET,
    cachePlugin,
  });
  await auth.initTenantFromUrl(url);

  const res = await auth.doAuthenticate(true);
  console.log(chalk`{yellow \nLogged in:}`);
  console.log(chalk`{yellow   user:} {green ${res.account.username}}`);
  console.log(chalk`{yellow tenant:} {green ${res.account.tenantId}}`);
  console.log(chalk`{yellow    exp:} {green ${res.expiresOn}}\n`);

  const client = new OneDrive({
    auth,
    noShareLinkCache: true,
  });
  return client;
}

async function run() {
  const contentBusId = process.argv[2];
  const url = process.argv[3];
  if (!url) {
    console.error('usage test-pwd-rotation <contentBusId> <mount-url>');
    process.exit(1);
  }
  const key = `${contentBusId}/.helix-auth/auth-onedrive-content.json`;
  const basePlugin = new S3CachePlugin({
    bucket: 'helix-content-bus',
    key,
    secret: contentBusId,
  });

  const cachePlugin = new MemCachePlugin({ key, base: basePlugin });

  let client = await getClient(cachePlugin, url);
  await testReadAccessOnedrive(client, url);

  await rl.question('\nRevoke user sessions in azure and press enter...');

  await clearAccessToken(basePlugin, true);

  client = await getClient(cachePlugin, url);
  await testReadAccessOnedrive(client, url);
}

run().catch(console.error).finally(() => {
  rl.close();
});
