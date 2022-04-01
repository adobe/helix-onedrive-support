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
import {AuthenticationResult, ClientApplication, ICachePlugin} from "@azure/msal-node";

export declare interface OneDriveAuthOptions {
  clientId: string;
  clientSecret?: string;
  refreshToken?: string;
  log?: Console;
  tenant?: string;
  resource?: string;
  scopes?: string[];
  localAuthCache?:boolean;

  /**
   * use cache plugin instead for default (global) token cache.
   */
  cachePlugin?: ICachePlugin,

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
  isAuthenticated(): boolean;

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
   * Performs a login using an interactive flow which prompts the user to open a browser window and
   * enter the authorization code.
   * @params {function} [onCode] - optional function that gets invoked after code was retrieved.
   * @returns {Promise<AuthenticationResult>}
   */
  acquireTokenByDeviceCode(onCode: Function): Promise<AuthenticationResult>;

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
  getAccessToken(silentOnly: boolean): Promise<AuthenticationResult>;

  dispose() : Promise<void>;

  initTenantFromUrl(sharingUrl: string|URL): Promise<void>;


}
