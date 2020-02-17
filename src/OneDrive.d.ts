/*
 * Copyright 2019 Adobe. All rights reserved.
 * This file is licensed to you under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License. You may obtain a copy
 * of the License at http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under
 * the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
 * OF ANY KIND, either express or implied. See the License for the specific language
 * governing permissions and limitations under the License.
 */
import { EventEmitter } from 'events';

/**
 * Logger interface
 */
declare interface Logger {
}

export declare interface OneDriveOptions {
  clientId: string;
  clientSecret?: string;
  refreshToken?: string;
  log?: Logger;
  accessToken?: string;
  expiresOn?: number;
  tenant?: string;
}

export declare interface GraphResult {
}

export declare interface DriveItem {
}

/**
 * Helper class that facilitates accessing one drive.
 */
export declare class OneDrive extends EventEmitter {
  /**
   * the maximum subscription time in milliseconds
   * @see https://docs.microsoft.com/en-us/graph/api/resources/subscription?view=graph-rest-1.0#maximum-length-of-subscription-per-resource-type
   */
  static MAX_SUBSCRIPTION_EXPIRATION_TIME: number;

  /**
   * Encodes the sharing url into a token that can be used to access drive items.
   * @param {string} sharingUrl A sharing URL from OneDrive
   * @see https://docs.microsoft.com/en-us/onedrive/developer/rest-api/api/shares_get?view=odsp-graph-online#encoding-sharing-urls
   * @returns {string} an id for a shared item.
   */
  static encodeSharingUrl(sharingUrl: string): string;

  /**
   * Creates a new OneDrive helper.
   * @param {OneDriveOptions} opts Options.
   */
  constructor(opts: OneDriveOptions);

  /**
   * is set to {@code true} if this client is initialized.
   */
  authenticated: boolean;

  /**
   * the logger of this client
   */
  log: Logger;

  /**
   * the authority url for login.
   */
  authorityUrl: string;

  /**
   * Performs a login using an interactive flow which prompts the user to open a browser window and
   * enter the authorization code.
   * @params {function} [onCode] - optional function that gets invoked after code was retrieved.
   * @returns {Promise<void>}
   */
  login(onCode: Function): Promise<any>;

  getAccessToken(autoRefresh: boolean): Promise<string>;

  createLoginUrl(): string;

  acquireToken(redirectUri: string, code: string): Promise<void>;

  getClient(): Promise<Request>;

  me(): Promise<GraphResult>;

  resolveShareLink(sharingUrl: string): Promise<GraphResult>;

  getDriveItemFromShareLink(sharingUrl: string): Promise<GraphResult>;

  listChildren(folderItem: DriveItem, relPath?: string): Promise<GraphResult>;

  getDriveItem(folderItem: DriveItem, relPath: string, download?: boolean): Promise<GraphResult>;

  downloadDriveItem(driveItem: DriveItem): Promise<GraphResult>;

  listSubscriptions(): Promise<GraphResult>;

  refreshSubscription(id: string, expiresIn: number): Promise<GraphResult>;

  uploadDriveItem(buffer: Buffer, driveItem: string, relPath?: string);

  /**
   * Returns the root item for a drive given its id.
   * @param driveId drive id
   * @returns {Promise<GraphResult>}
   */
  getDriveRootItem(driveId: string): Promise<GraphResult>;

  /**
   * Fetches the changes from the respective resource using the provided delta token.
   * Use an empty token to fetch the initial state or `latest` to fetch the latest state.
   * @param {string} resource OneDrive resource path.
   * @param {string} [token] Delta token.
   * @returns {Promise<Array>} A return object with the values and a `@odata.deltaLink`.
   */
  fetchChanges(resource: string, token?: string);
}
