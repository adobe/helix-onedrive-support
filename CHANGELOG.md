# [2.8.0](https://github.com/adobe/helix-onedrive-support/compare/v2.7.0...v2.8.0) (2020-07-17)


### Features

* **onedrive:** add support for onedrive uri ([8ec07ec](https://github.com/adobe/helix-onedrive-support/commit/8ec07ec74ee49ebb67588df0ac2d455a82449021)), closes [#68](https://github.com/adobe/helix-onedrive-support/issues/68) [#69](https://github.com/adobe/helix-onedrive-support/issues/69)

# [2.7.0](https://github.com/adobe/helix-onedrive-support/compare/v2.6.0...v2.7.0) (2020-07-15)


### Features

* **test:** add MockOneDrive for emulating excel workbooks ([5995ca2](https://github.com/adobe/helix-onedrive-support/commit/5995ca28aaf61ecdc14920e6baacc01d3009acdb))

# [2.6.0](https://github.com/adobe/helix-onedrive-support/compare/v2.5.1...v2.6.0) (2020-07-14)


### Features

* **excel:** add support for ranges and object serialization ([#66](https://github.com/adobe/helix-onedrive-support/issues/66)) ([2220028](https://github.com/adobe/helix-onedrive-support/commit/2220028d0458ce04a93984ef75696ccddc931f05))

## [2.5.1](https://github.com/adobe/helix-onedrive-support/compare/v2.5.0...v2.5.1) (2020-06-27)


### Bug Fixes

* **table:** support addRows ([4eec5fd](https://github.com/adobe/helix-onedrive-support/commit/4eec5fda1a9d97d8ff547788dce6f3c1900187ec))

# [2.5.0](https://github.com/adobe/helix-onedrive-support/compare/v2.4.1...v2.5.0) (2020-06-16)


### Features

* **table:** add getColumn method and simplify row handling ([#55](https://github.com/adobe/helix-onedrive-support/issues/55)) ([f02c440](https://github.com/adobe/helix-onedrive-support/commit/f02c4406ed1775c2c1734bfc0586b79bf47443e5))

## [2.4.1](https://github.com/adobe/helix-onedrive-support/compare/v2.4.0...v2.4.1) (2020-06-15)


### Bug Fixes

* **table:** add row count ([#48](https://github.com/adobe/helix-onedrive-support/issues/48)) ([f0b8e7e](https://github.com/adobe/helix-onedrive-support/commit/f0b8e7e2ec9f8467b5203afbbbd896d4eb0c7851))

# [2.4.0](https://github.com/adobe/helix-onedrive-support/compare/v2.3.0...v2.4.0) (2020-05-28)


### Features

* **excel:** support Workbooks and Tables ([#44](https://github.com/adobe/helix-onedrive-support/issues/44)) ([a285651](https://github.com/adobe/helix-onedrive-support/commit/a285651930afac700ebabab126dc8cb1d86075a1))

# [2.3.0](https://github.com/adobe/helix-onedrive-support/compare/v2.2.0...v2.3.0) (2020-04-16)


### Features

* **subscription:** provide create/delete of subscriptions ([#35](https://github.com/adobe/helix-onedrive-support/issues/35)) ([29094d5](https://github.com/adobe/helix-onedrive-support/commit/29094d5a2c148837294d90520d051843b64b8d13))

# [2.2.0](https://github.com/adobe/helix-onedrive-support/compare/v2.1.0...v2.2.0) (2020-03-26)


### Features

* **cache:** improve token cache ([#32](https://github.com/adobe/helix-onedrive-support/issues/32)) ([ad59d0f](https://github.com/adobe/helix-onedrive-support/commit/ad59d0f6769c742d8aa17c9682fbdcd9f78c95b0)), closes [#31](https://github.com/adobe/helix-onedrive-support/issues/31)

# [2.1.0](https://github.com/adobe/helix-onedrive-support/compare/v2.0.0...v2.1.0) (2020-02-21)


### Features

* **auth:** add support for username+password authentication ([#27](https://github.com/adobe/helix-onedrive-support/issues/27)) ([a1eb863](https://github.com/adobe/helix-onedrive-support/commit/a1eb863d93166200dcf12bf7b8dcbff79a8f78e8))

# [2.0.0](https://github.com/adobe/helix-onedrive-support/compare/v1.4.1...v2.0.0) (2020-02-19)


### Bug Fixes

* rename API ([#25](https://github.com/adobe/helix-onedrive-support/issues/25)) ([7c1cffa](https://github.com/adobe/helix-onedrive-support/commit/7c1cffa6b198246d6abb838a495b206a47a2b5de))


### BREAKING CHANGES

* renamed API to getDriveRootItem()
* return token from fetchChanges (fix #26)

## [1.4.1](https://github.com/adobe/helix-onedrive-support/compare/v1.4.0...v1.4.1) (2020-02-18)


### Bug Fixes

* **deps:** update external ([#24](https://github.com/adobe/helix-onedrive-support/issues/24)) ([8ff122f](https://github.com/adobe/helix-onedrive-support/commit/8ff122f853244391c6e6a3d73c137da71c39b11e))

# [1.4.0](https://github.com/adobe/helix-onedrive-support/compare/v1.3.1...v1.4.0) (2020-02-07)


### Features

* **rootfolder:** add support for getting the root folder id ([#22](https://github.com/adobe/helix-onedrive-support/issues/22)) ([40c682f](https://github.com/adobe/helix-onedrive-support/commit/40c682f9592921039236dc265367c9823ac451f2))

## [1.3.1](https://github.com/adobe/helix-onedrive-support/compare/v1.3.0...v1.3.1) (2020-02-06)


### Bug Fixes

* **pur:** Put should not be set to json ([#21](https://github.com/adobe/helix-onedrive-support/issues/21)) ([2f61840](https://github.com/adobe/helix-onedrive-support/commit/2f61840b038c6fa77209ce88ffc025c5006ce30e))

# [1.3.0](https://github.com/adobe/helix-onedrive-support/compare/v1.2.1...v1.3.0) (2020-02-05)


### Features

* **login:** provide callback instead of flag ([#19](https://github.com/adobe/helix-onedrive-support/issues/19)) ([5d255e9](https://github.com/adobe/helix-onedrive-support/commit/5d255e9939a98a1951cba424c5b4d69232acb25b))

## [1.2.1](https://github.com/adobe/helix-onedrive-support/compare/v1.2.0...v1.2.1) (2020-02-04)


### Bug Fixes

* **core:** client secret must not be mandatory ([#18](https://github.com/adobe/helix-onedrive-support/issues/18)) ([ba1945c](https://github.com/adobe/helix-onedrive-support/commit/ba1945c7b74ff418d57e0732fb7d4c7c9cffbd83))

# [1.2.0](https://github.com/adobe/helix-onedrive-support/compare/v1.1.0...v1.2.0) (2020-02-04)


### Features

* **auth:** add support for interactive login flow ([#17](https://github.com/adobe/helix-onedrive-support/issues/17)) ([2d19e33](https://github.com/adobe/helix-onedrive-support/commit/2d19e3386fb336794b21cd21c5638cdff1f4d992)), closes [#16](https://github.com/adobe/helix-onedrive-support/issues/16)

# [1.1.0](https://github.com/adobe/helix-onedrive-support/compare/v1.0.4...v1.1.0) (2020-02-03)


### Features

* **core:** add support for uploading drive items ([#12](https://github.com/adobe/helix-onedrive-support/issues/12)) ([37665a7](https://github.com/adobe/helix-onedrive-support/commit/37665a7755b0379c81e496c94a5570f3b498f94c))

## [1.0.4](https://github.com/adobe/helix-onedrive-support/compare/v1.0.3...v1.0.4) (2020-01-29)


### Bug Fixes

* **api:** make relpath optional in list children ([#11](https://github.com/adobe/helix-onedrive-support/issues/11)) ([6fef240](https://github.com/adobe/helix-onedrive-support/commit/6fef240169aa13be886eab59b3f5b3fa36fcb95f))

## [1.0.3](https://github.com/adobe/helix-onedrive-support/compare/v1.0.2...v1.0.3) (2020-01-27)


### Bug Fixes

* **client:** set access token correctly ([#9](https://github.com/adobe/helix-onedrive-support/issues/9)) ([92464b5](https://github.com/adobe/helix-onedrive-support/commit/92464b5f753f4ef5b38cd77fb6ded50d58867e74))

## [1.0.2](https://github.com/adobe/helix-onedrive-support/compare/v1.0.1...v1.0.2) (2020-01-24)


### Bug Fixes

* **ci:** trigger release ([3dc55c6](https://github.com/adobe/helix-onedrive-support/commit/3dc55c626d5e26193fd6210ac603b3cf94ee1465))

## [1.0.1](https://github.com/adobe/helix-onedrive-support/compare/v1.0.0...v1.0.1) (2020-01-24)


### Bug Fixes

* check for required arguments clientId and clientSecret ([5b3e8d1](https://github.com/adobe/helix-onedrive-support/commit/5b3e8d16c620007359ef2c223079e693360e0032))
* **test:** add test for required parameters ([10cf10d](https://github.com/adobe/helix-onedrive-support/commit/10cf10d935d97333d0ae3f848fdfc3062c83635e))
* **test:** remove try/catch block ([73498c5](https://github.com/adobe/helix-onedrive-support/commit/73498c55e53dafc7f0321d342bad1eddd8c859ee))
* **test:** use assert.throws() ([94b5989](https://github.com/adobe/helix-onedrive-support/commit/94b598978b26567b63ae9506304cad774297ce48))

# 1.0.0 (2020-01-17)


### Bug Fixes

* **core:** initial release ([34baa59](https://github.com/adobe/helix-onedrive-support/commit/34baa59f2207ba51db28543022e86846a7432d86))
