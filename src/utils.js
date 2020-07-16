/*
 * Copyright 2020 Adobe. All rights reserved.
 * This file is licensed to you under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License. You may obtain a copy
 * of the License at http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under
 * the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
 * OF ANY KIND, either express or implied. See the License for the specific language
 * governing permissions and limitations under the License.
 */

/**
 * Returns a onedrive uri for the given drive item. the uri has the format:
 * `onedrive:/drives/<driveId>/items/<itemId>`
 *
 * @param {DriveItem} driveItem
 * @returns {URL} An url representing the drive item
 */
function driveItemToURL(driveItem) {
  return new URL(`onedrive:/drives/${driveItem.parentReference.driveId}/items/${driveItem.id}`);
}

/**
 * Returns a partial drive item from the given url. The urls needs to have the format:
 * `onedrive:/drives/<driveId>/items/<itemId>`. if the url does not start with the correct
 * protocol, {@code null} is returned.
 *
 * @param {URL|string} url The url of the drive item.
 * @return {DriveItem} A (partial) drive item.
 */
function driveItemFromURL(url) {
  if (!(url instanceof URL)) {
    // eslint-disable-next-line no-param-reassign
    url = new URL(String(url));
  }
  if (url.protocol !== 'onedrive:') {
    return null;
  }
  const [drives, driveId, items, itemId] = url.pathname.split('/').filter((s) => !!s);
  if (drives !== 'drives') {
    throw new Error(`URI not supported (missing 'drives' segment): ${url}`);
  }
  if (items !== 'items') {
    throw new Error(`URI not supported (missing 'items' segment): ${url}`);
  }
  return {
    id: itemId,
    parentReference: {
      driveId,
    },
  };
}

module.exports = {
  driveItemFromURL,
  driveItemToURL,
};
