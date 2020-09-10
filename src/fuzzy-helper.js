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
 * Splits the given name at the last '.', returning the extension and the base name.
 * @param {string} name Filename
 * @returns {string[]} Returns an array containing the base name and extension.
 */
function splitByExtension(name) {
  const idx = name.lastIndexOf('.');
  const baseName = idx > 0 && idx < name.length - 1 ? name.substring(0, idx) : name;
  const ext = idx > 0 && idx < name.length - 1 ? name.substring(idx + 1).toLowerCase() : '';
  return [baseName, ext];
}

/**
 * Sanitizes the given string by :
 * - convert to lower case
 * - replace all non-alphanumeric characters with a dash
 * - remove all consecutive dashes
 *
 * @param {string} name
 * @returns {string} sanitized name
 */
function sanitize(name) {
  return name.toLowerCase().replace(/[^a-z0-9]+/g, '-');
}

/**
 * Compute the edit distance using a recursive algorithm. since we only expect to have relative
 * short filenames, the algorithm shouldn't be too expensive.
 *
 * @param {string} s0 Input string
 * @param {string} s1 Input string
 * @returns {number|*}
 */
function editDistance(s0, s1) {
  // make sure that s0 length is greater than s1 length
  if (s0.length < s1.length) {
    const t = s1;
    // eslint-disable-next-line no-param-reassign
    s1 = s0;
    // eslint-disable-next-line no-param-reassign
    s0 = t;
  }
  const l0 = s0.length;
  const l1 = s1.length;

  // init first row
  const resultMatrix = [[]];
  for (let c = 0; c < l1 + 1; c += 1) {
    resultMatrix[0][c] = c;
  }
  // fill out the distance matrix and find the best path
  for (let i = 1; i < l0 + 1; i += 1) {
    resultMatrix[i] = [i];
    for (let j = 1; j < l1 + 1; j += 1) {
      const replaceCost = (s0.charAt(i - 1) === s1.charAt(j - 1)) ? 0 : 1;
      resultMatrix[i][j] = Math.min(
        resultMatrix[i - 1][j] + 1, // insert
        resultMatrix[i][j - 1] + 1, // remove
        resultMatrix[i - 1][j - 1] + replaceCost,
      );
    }
  }
  return resultMatrix[l0][l1];
}

module.exports = {
  splitByExtension,
  sanitize,
  editDistance,
};
