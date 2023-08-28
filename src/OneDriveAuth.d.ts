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
import {AuthenticationResult, ClientApplication } from "@azure/msal-node";
import { CachePlugin } from "@adobe/helix-shared-tokencache";

export enum AcquireMethod {
  BY_CLIENT_CREDENTIAL = 'byClientCredential',
  BY_DEVICE_CODE = 'byDeviceCode',
}

export declare interface OneDriveAuthOptions {
  clientId?: string;
  clientSecret?: string;
  log?: Console;
  tenant?: string;
  scopes?: string[];
  onCode?: Function;
  localAuthCache?:boolean;
  acquireMethod?: AcquireMethod;
  accessToken?: string;

  /**
   * Optional log fields, as key-value object.
   */
  logFields?: object;

  /**
   * use cache plugin instead for default (global) token cache.
   */
  cachePlugin?: CachePlugin,

  /**
   * Disables the cache for the tenant lookup.
   * @default process.env.HELIX_ONEDRIVE_NO_TENANT_CACHE
   */
  noTenantCache?: boolean;

  /**
   * Map to use for the tenant lookup cache. If empty, a module-global cache will be used.
   * Note that the cache is only used, if the `noTenantCache` flag is `falsy`
   */
  tenantCache?: Map<string, string>;
}

/**
 * Helix config mount point
 */
declare interface MountPoint {
  url: string;
  tenantId?: string;
}

/**
 * Helper class that facilitates authentication for one drive.
 */
export declare class OneDriveAuth {
  /**
   * Creates a new OneDriveAuth helper.
   * @param {OneDriveAuthOptions} opts Options.
   */
  constructor(opts: OneDriveAuthOptions);

  /**
   * {@code true} if this client is initialized.
   */
  isAuthenticated(): Promise<boolean>;

  /**
   * the logger of this client
   */
  log: Console;

  /**
   * the MSAL client application
   */
  app: ClientApplication;

  /**
   * the authority url for login.
   */
  getAuthorityUrl(): string;

  /**
   * Sets the access token to use for all requests. if the token is a valid JWT token,
   * its `tid` claim is used a tenant (if no tenant is already set).
   *
   * @param {string} bearerToken
   * @returns {OneDriveAuth} this
   */
  setAccessToken(bearerToken): OneDriveAuth;

  /**
   * Acquires the access token either from the cache or from MS.
   * @param {boolean} silentOnly
   * @returns {Promise<AuthenticationResult>}
   */
  authenticate(silentOnly: boolean): Promise<AuthenticationResult>;

  /**
   * Disposes allocated resources.
   */
  dispose() : Promise<void>;

  /**
   * Initializes the tenant from the url by requesting the required information from microsoft.
   * @param {string|URL} sharingUrl
   * @returns {string} the tenant id
   */
  initTenantFromUrl(sharingUrl: string|URL): Promise<string>;

  /**
   * Initializes the tenant either with the `tenantId` of the mount point or via the share url.
   * @param {MountPoint} mp
   * @returns {string} the tenant id
   */
  initTenantFromMountPoint(mp: MountPoint): Promise<string>;
}
