# Helix OneDrive Support

> An example library to be used in and with Project Helix

## Status
[![codecov](https://img.shields.io/codecov/c/github/adobe/helix-onedrive-support.svg)](https://codecov.io/gh/adobe/helix-onedrive-support)
[![CircleCI](https://img.shields.io/circleci/project/github/adobe/helix-onedrive-support.svg)](https://circleci.com/gh/adobe/helix-onedrive-support)
[![GitHub license](https://img.shields.io/github/license/adobe/helix-onedrive-support.svg)](https://github.com/adobe/helix-onedrive-support/blob/master/LICENSE.txt)
[![GitHub issues](https://img.shields.io/github/issues/adobe/helix-onedrive-support.svg)](https://github.com/adobe/helix-onedrive-support/issues)
[![LGTM Code Quality Grade: JavaScript](https://img.shields.io/lgtm/grade/javascript/g/adobe/helix-onedrive-support.svg?logo=lgtm&logoWidth=18)](https://lgtm.com/projects/g/adobe/helix-onedrive-support)
[![semantic-release](https://img.shields.io/badge/%20%20%F0%9F%93%A6%F0%9F%9A%80-semantic--release-e10079.svg)](https://github.com/semantic-release/semantic-release)

## Installation

```bash
$ npm install @adobe/helix-onedrive-support
```

## Usage

See the [API documentation](docs/API.md).

## Development

### Build

```bash
$ npm install
```

### Test

```bash
$ npm test
```

### Lint

```bash
$ npm run lint
```

## Testing

You can browse the OneDrive integration using _browser.js_:

1. start with: `npm start`.
2. open web browser at `http://localhost:3000/`.
3. sign in with Microsoft.
3. copy-paste share-link of a shared folder (see below) and click the `list` button.


## Authentication

The action authenticates against OneDrive using an oauth2 refresh token. In case you need to
create a new one, use the _browser.js_ to generate a `tokens.json`.
