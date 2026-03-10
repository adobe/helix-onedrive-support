/*
 * Copyright 2026 Adobe. All rights reserved.
 * This file is licensed to you under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License. You may obtain a copy
 * of the License at http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under
 * the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
 * OF ANY KIND, either express or implied. See the License for the specific language
 * governing permissions and limitations under the License.
 */
import { Headers } from '@adobe/fetch';
import { createAuthError, ClientAuthErrorCodes, createNetworkError } from '@azure/msal-common/node';

/**
 * HTTP methods
 */
const HttpMethod = {
  GET: 'GET',
  POST: 'POST',
};

/**
 * Converts a fetch Headers object to a plain JavaScript object.
 *
 * The fetch API returns headers as a Headers object with methods like get(), has(),
 * etc. However, the rest of the MSAL codebase expects headers as a simple key-value
 * object. This function performs that conversion.
 *
 * @param headers - The Headers object returned by fetch response
 * @returns A plain object with header names as keys and values as strings
 */
function getHeaderDict(headers) {
  const headerDict = {};
  headers.forEach((value, key) => {
    headerDict[key] = value;
  });
  return headerDict;
}
/**
 * Converts NetworkRequestOptions headers to a fetch-compatible Headers object.
 *
 * The MSAL library uses plain objects for headers in NetworkRequestOptions,
 * but the fetch API expects either a Headers object, plain object, or array
 * of arrays. Using the Headers constructor provides better compatibility
 * and validation.
 *
 * @param options - Optional NetworkRequestOptions containing headers
 * @returns A Headers object ready for use with fetch API
 */
function getFetchHeaders(options) {
  const headers = new Headers();
  if (options?.headers) {
    Object.entries(options.headers).forEach(([key, value]) => {
      headers.append(key, value);
    });
  }
  return headers;
}

/**
 * Custom network module implementation using @adobe/fetch.
 */
export class CustomNetworkModule {
  /**
   * @constructor
   * @param {import('@adobe/fetch').FetchContext} fetchContext fetch context
   */
  constructor(fetchContext) {
    this.fetchContext = fetchContext;
  }

  /**
   * Sends an HTTP GET request to the specified URL.
   *
   * This method handles GET requests with optional timeout support. The timeout
   * is implemented using AbortController, which provides a clean way to cancel
   * fetch requests that take too long to complete.
   *
   * @param url - The target URL for the GET request
   * @param options - Optional request configuration including headers
   * @param timeout - Optional timeout in milliseconds. If specified, the request
   *                  will be aborted if it doesn't complete within this time
   * @returns Promise that resolves to a NetworkResponse containing headers, body, and status
   * @throws {AuthError} When the request times out or response parsing fails
   * @throws {NetworkError} When the network request fails
   */
  async sendGetRequestAsync(url, options, timeout) {
    return this.sendRequest(url, HttpMethod.GET, options, timeout);
  }

  /**
   * Sends an HTTP POST request to the specified URL.
   *
   * This method handles POST requests with request body support. Currently,
   * timeout functionality is not exposed for POST requests, but the underlying
   * implementation supports it through the shared sendRequest method.
   *
   * @param url - The target URL for the POST request
   * @param options - Optional request configuration including headers and body
   * @returns Promise that resolves to a NetworkResponse containing headers, body, and status
   * @throws {AuthError} When the request times out or response parsing fails
   * @throws {NetworkError} When the network request fails
   */
  async sendPostRequestAsync(url, options) {
    return this.sendRequest(url, HttpMethod.POST, options);
  }

  /**
   * Send an HTTP request to the specified URL.
   *
   * @param url - The target URL for the request
   * @param method - HTTP method (GET or POST)
   * @param options - Optional request configuration (headers, body)
   * @param timeout - Optional timeout in milliseconds for request cancellation
   * @returns Promise resolving to NetworkResponse with parsed JSON body
   * @throws {AuthError} For timeouts or JSON parsing errors
   * @throws {NetworkError} For network failures
   */
  async sendRequest(url, method, options, timeout) {
    const fetchOptions = {
      method,
      headers: getFetchHeaders(options),
    };
    /*
     * Configure timeout if specified
     * The setTimeout will trigger abort() if the request takes too long
     */
    let timerId;
    if (timeout) {
      const controller = new AbortController();
      timerId = setTimeout(() => controller.abort(), timeout);
      fetchOptions.signal = controller.signal;
    }
    if (method === HttpMethod.POST) {
      fetchOptions.body = options?.body || '';
    }

    let response;
    try {
      const { fetch } = this.fetchContext;
      response = await fetch(url, fetchOptions);
    } catch (error) {
      // Clean up timeout to prevent memory leaks
      if (timerId) {
        clearTimeout(timerId);
      }
      if (error instanceof Error && error.name === 'AbortError') {
        throw createAuthError(ClientAuthErrorCodes.networkError, 'Request timeout');
      }
      const baseAuthError = createAuthError(ClientAuthErrorCodes.networkError, `Network request failed: ${error instanceof Error ? error.message : /* c8 ignore next */ 'unknown'}`);
      throw createNetworkError(
        baseAuthError,
        undefined,
        undefined,
        error instanceof Error ? error : /* c8 ignore next */ undefined,
      );
    }
    try {
      return {
        headers: getHeaderDict(response.headers),
        body: (await response.json()),
        status: response.status,
      };
    } catch (error) {
      throw createAuthError(ClientAuthErrorCodes.tokenParsingError, `Failed to parse response: ${error instanceof Error ? error.message : /* c8 ignore next */ 'unknown'}`);
    }
  }
}
