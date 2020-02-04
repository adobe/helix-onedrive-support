<a name="OneDrive"></a>

## OneDrive
Helper class that facilitates accessing one drive.

**Kind**: global class  

* [OneDrive](#OneDrive)
    * [new OneDrive(opts)](#new_OneDrive_new)
    * _instance_
        * [.log](#OneDrive+log)
        * [.authenticated](#OneDrive+authenticated) ⇒ <code>boolean</code>
        * [.login()](#OneDrive+login) ⇒ <code>Promise.&lt;void&gt;</code>
        * [.getAccessToken()](#OneDrive+getAccessToken)
        * [.createLoginUrl()](#OneDrive+createLoginUrl)
        * [.acquireToken()](#OneDrive+acquireToken)
        * [.getClient()](#OneDrive+getClient)
        * [.resolveShareLink()](#OneDrive+resolveShareLink)
        * [.getDriveItemFromShareLink()](#OneDrive+getDriveItemFromShareLink)
        * [.listChildren()](#OneDrive+listChildren)
        * [.getDriveItem()](#OneDrive+getDriveItem)
        * [.downloadDriveItem()](#OneDrive+downloadDriveItem)
        * [.uploadDriveItem()](#OneDrive+uploadDriveItem)
        * [.listSubscriptions()](#OneDrive+listSubscriptions)
        * [.refreshSubscription()](#OneDrive+refreshSubscription)
        * [.fetchChanges(resource, [token])](#OneDrive+fetchChanges) ⇒ <code>Promise.&lt;Array&gt;</code>
    * _static_
        * [.MAX_SUBSCRIPTION_EXPIRATION_TIME](#OneDrive.MAX_SUBSCRIPTION_EXPIRATION_TIME)
        * [.encodeSharingUrl(sharingUrl)](#OneDrive.encodeSharingUrl) ⇒ <code>string</code>

<a name="new_OneDrive_new"></a>

### new OneDrive(opts)

| Param | Type | Description |
| --- | --- | --- |
| opts | <code>OneDriveOptions</code> | Options |
| opts.clientId | <code>string</code> | The client id of the app |
| opts.clientSecret | <code>string</code> | The client secret of the app |
| [opts.refreshToken] | <code>string</code> | The refresh token. |
| [opts.refreshToken] | <code>string</code> | The access token. |
| [opts.expiresOn] | <code>number</code> | Expiration time. |
| [opts.log] | <code>Logger</code> | A logger. |

<a name="OneDrive+log"></a>

### oneDrive.log
**Kind**: instance property of [<code>OneDrive</code>](#OneDrive)  
<a name="OneDrive+authenticated"></a>

### oneDrive.authenticated ⇒ <code>boolean</code>
**Kind**: instance property of [<code>OneDrive</code>](#OneDrive)  
<a name="OneDrive+login"></a>

### oneDrive.login() ⇒ <code>Promise.&lt;void&gt;</code>
Performs a login using an interactive flow which prompts the user to open a browser window and
enter the authorization code.

**Kind**: instance method of [<code>OneDrive</code>](#OneDrive)  
**Params**: <code>boolean</code> open - if true, automatically opens the default browser  
<a name="OneDrive+getAccessToken"></a>

### oneDrive.getAccessToken()
**Kind**: instance method of [<code>OneDrive</code>](#OneDrive)  
<a name="OneDrive+createLoginUrl"></a>

### oneDrive.createLoginUrl()
**Kind**: instance method of [<code>OneDrive</code>](#OneDrive)  
<a name="OneDrive+acquireToken"></a>

### oneDrive.acquireToken()
**Kind**: instance method of [<code>OneDrive</code>](#OneDrive)  
<a name="OneDrive+getClient"></a>

### oneDrive.getClient()
**Kind**: instance method of [<code>OneDrive</code>](#OneDrive)  
<a name="OneDrive+resolveShareLink"></a>

### oneDrive.resolveShareLink()
**Kind**: instance method of [<code>OneDrive</code>](#OneDrive)  
<a name="OneDrive+getDriveItemFromShareLink"></a>

### oneDrive.getDriveItemFromShareLink()
**Kind**: instance method of [<code>OneDrive</code>](#OneDrive)  
<a name="OneDrive+listChildren"></a>

### oneDrive.listChildren()
**Kind**: instance method of [<code>OneDrive</code>](#OneDrive)  
<a name="OneDrive+getDriveItem"></a>

### oneDrive.getDriveItem()
**Kind**: instance method of [<code>OneDrive</code>](#OneDrive)  
<a name="OneDrive+downloadDriveItem"></a>

### oneDrive.downloadDriveItem()
**Kind**: instance method of [<code>OneDrive</code>](#OneDrive)  
<a name="OneDrive+uploadDriveItem"></a>

### oneDrive.uploadDriveItem()
**Kind**: instance method of [<code>OneDrive</code>](#OneDrive)  
**See**: https://docs.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0&tabs=http  
<a name="OneDrive+listSubscriptions"></a>

### oneDrive.listSubscriptions()
**Kind**: instance method of [<code>OneDrive</code>](#OneDrive)  
<a name="OneDrive+refreshSubscription"></a>

### oneDrive.refreshSubscription()
**Kind**: instance method of [<code>OneDrive</code>](#OneDrive)  
<a name="OneDrive+fetchChanges"></a>

### oneDrive.fetchChanges(resource, [token]) ⇒ <code>Promise.&lt;Array&gt;</code>
Fetches the changes from the respective resource using the provided delta token.
Use an empty token to fetch the initial state or `latest` to fetch the latest state.

**Kind**: instance method of [<code>OneDrive</code>](#OneDrive)  
**Returns**: <code>Promise.&lt;Array&gt;</code> - A return object with the values and a `@odata.deltaLink`.  

| Param | Type | Description |
| --- | --- | --- |
| resource | <code>string</code> | OneDrive resource path. |
| [token] | <code>string</code> | Delta token. |

<a name="OneDrive.MAX_SUBSCRIPTION_EXPIRATION_TIME"></a>

### OneDrive.MAX\_SUBSCRIPTION\_EXPIRATION\_TIME
the maximum subscription time in milliseconds

**Kind**: static constant of [<code>OneDrive</code>](#OneDrive)  
**See**: https://docs.microsoft.com/en-us/graph/api/resources/subscription?view=graph-rest-1.0#maximum-length-of-subscription-per-resource-type  
<a name="OneDrive.encodeSharingUrl"></a>

### OneDrive.encodeSharingUrl(sharingUrl) ⇒ <code>string</code>
Encodes the sharing url into a token that can be used to access drive items.

**Kind**: static method of [<code>OneDrive</code>](#OneDrive)  
**Returns**: <code>string</code> - an id for a shared item.  
**See**: https://docs.microsoft.com/en-us/onedrive/developer/rest-api/api/shares_get?view=odsp-graph-online#encoding-sharing-urls  

| Param | Type | Description |
| --- | --- | --- |
| sharingUrl | <code>string</code> | A sharing URL from OneDrive |

