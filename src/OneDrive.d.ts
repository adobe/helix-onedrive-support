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
import { Workbook } from './Workbook';
import { TokenResponse } from 'adal-node';

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
  tenant?: string;
  username?: string;
  password?: string;
}

export declare interface GraphResult {
}

export declare interface DriveItem {
}

export declare interface SubscriptionOptions {
  resource: string;
  notificationUrl: string;
  clientState: string;
  changeType?: string;
  expiresIn?: number;
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
   * @param {string|URL} sharingUrl A sharing URL from OneDrive
   * @see https://docs.microsoft.com/en-us/onedrive/developer/rest-api/api/shares_get?view=odsp-graph-online#encoding-sharing-urls
   * @returns {string} an id for a shared item.
   */
  static encodeSharingUrl(sharingUrl: string|URL): string;

  /**
   * Returns a onedrive uri for the given drive item. the uri has the format:
   * `onedrive:/drives/<driveId>/items/<itemId>`
   *
   * @param {DriveItem} driveItem
   * @returns {URL} An url representing the drive item
   */
  static driveItemToURL(driveItem: DriveItem): URL;

  /**
   * Returns a partial drive item from the given url. The urls needs to have the format:
   * `onedrive:/drives/<driveId>/items/<itemId>`
   *
   * @param {URL|string} url The url of the drive item.
   * @return {DriveItem} A (partial) drive item.
   */
  static driveItemFromURL(url: URL): DriveItem;

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
   * Adds entries to the token cache
   * @param {TokenResponse[]} entries
   */
  loadTokenCache(entries: TokenResponse[]): Promise<void>;

  /**
   * Performs a login using an interactive flow which prompts the user to open a browser window and
   * enter the authorization code.
   * @params {function} [onCode] - optional function that gets invoked after code was retrieved.
   * @returns {Promise<TokenResponse>}
   */
  login(onCode: Function): Promise<TokenResponse>;

  getAccessToken(autoRefresh: boolean): Promise<TokenResponse>;

  createLoginUrl(): string;

  acquireToken(redirectUri: string, code: string): Promise<TokenResponse>;

  getClient(): Promise<Request>;

  me(): Promise<GraphResult>;

  resolveShareLink(sharingUrl: string|URL): Promise<GraphResult>;

  /**
   * Returns a drive item from the given share link or onedrive uri.
   * @param {string|URL} url The share link url or a onedrive uri.
   * @see OneDrive.driveItemToURL
   */
  getDriveItemFromShareLink(url: string|URL): Promise<DriveItem>;

  listChildren(folderItem: DriveItem, relPath?: string, query?: object): Promise<GraphResult>;

  /**
   * Returns the drive item for the given folder id and rel path.
   * If the relPath is empty, the folder item is returned.
   *
   * @param {DriveItem} folderItem Folder Item.
   * @param {string} [relPath=''] Relative path of item to retrieved
   * @param {boolean} [download=false] If {@code true} download the item instead.
   */
  getDriveItem(folderItem: DriveItem, relPath?: string, download?: boolean): Promise<GraphResult>;

  /**
   * Tries to get the drive items for the given folder and relative path, by loading the files of
   * the respective directory and returning the item with the best matching filename. Please note,
   * that only the files are matched 'fuzzily' but not the folders. The rules for transforming the
   * filenames to the name segment of the `relPath` are:
   * - convert to lower case
   * - normalize all unicode characters
   * - replace all non-alphanumeric characters with a dash
   * - remove all consecutive dashes
   * - remove all leading and trailing dashes
   * - extensions are ignored, if the given path doesn't have one
   *
   * The result is an array of drive items that match the given path. They are ordered by the edit
   * distance to the original name and then alphanumerically.
   *
   * @param folderItem
   * @param relPath
   * @returns {Promise<DriveItem[]>}
   */
  fuzzyGetDriveItem(folderItem: DriveItem, relPath?: string): Promise<DriveItem[]>;

  downloadDriveItem(driveItem: DriveItem): Promise<GraphResult>;

  /**
   * Creates a new workbook instance from a drive item.
   *
   * @param {DriveItem} fileItem drive item
   * @returns {Workbook} workbook instance
   */
  getWorkbook(fileItem: DriveItem): Workbook;

  listSubscriptions(): Promise<GraphResult>;

  createSubscription(opts: SubscriptionOptions): Promise<GraphResult>;

  refreshSubscription(id: string, expiresIn: number): Promise<GraphResult>;

  deleteSubscription(id: string): Promise<GraphResult>;

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
