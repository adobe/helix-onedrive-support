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
 * Trims the string at both ends and removes the zero width unicode chars:
 *
 * - U+200B zero width space
 * - U+200C zero width non-joiner Unicode code point
 * - U+200D zero width joiner Unicode code point
 * - U+FEFF zero width no-break space Unicode code point
 *
 * @param {string} str input string
 * @return {string} trimmed and stripped string
 */
export function superTrim(str) {
  return String(str)
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .trim();
}
