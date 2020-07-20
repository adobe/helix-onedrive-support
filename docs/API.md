## Classes

<dl>
<dt><a href="#OneDrive">OneDrive</a></dt>
<dd><p>Helper class that facilitates accessing one drive.</p>
</dd>
<dt><a href="#OneDriveMock">OneDriveMock</a></dt>
<dd><p>Mock OneDrive client that supports some of the operations the <code>OneDrive</code> class does.</p>
</dd>
</dl>

## Functions

<dl>
<dt><a href="#getActualError">getActualError(e)</a></dt>
<dd><p>Returns the actual error, recursively descending through all error properties.</p>
</dd>
<dt><a href="#handleNamedItems">handleNamedItems(sheet, segs, method, body)</a> ⇒ <code>object</code></dt>
<dd><p>Handle the <code>namedItems</code> operation on a workbook / worksheet</p>
</dd>
<dt><a href="#handleTable">handleTable(sheet, segs, method, body)</a> ⇒ <code>object</code></dt>
<dd><p>Handle the <code>table</code> operation on a workbook / worksheet</p>
</dd>
<dt><a href="#getActualError">getActualError(e)</a></dt>
<dd><p>Returns the actual error, recursively descending through all error properties.</p>
</dd>
<dt><a href="#getActualError">getActualError(e)</a></dt>
<dd><p>Returns the actual error, recursively descending through all error properties.</p>
</dd>
<dt><a href="#driveItemToURL">driveItemToURL(driveItem)</a> ⇒ <code>URL</code></dt>
<dd><p>Returns a onedrive uri for the given drive item. the uri has the format:
<code>onedrive:/drives/&lt;driveId&gt;/items/&lt;itemId&gt;</code></p>
</dd>
<dt><a href="#driveItemFromURL">driveItemFromURL(url)</a> ⇒ <code>DriveItem</code></dt>
<dd><p>Returns a partial drive item from the given url. The urls needs to have the format:
<code>onedrive:/drives/&lt;driveId&gt;/items/&lt;itemId&gt;</code>. if the url does not start with the correct
protocol, {@code null} is returned.</p>
</dd>
</dl>

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
        * [.getDriveRootItem()](#OneDrive+getDriveRootItem)
        * [.getDriveItemFromShareLink()](#OneDrive+getDriveItemFromShareLink)
        * [.listChildren()](#OneDrive+listChildren)
        * [.getDriveItem()](#OneDrive+getDriveItem)
        * [.downloadDriveItem()](#OneDrive+downloadDriveItem)
        * [.uploadDriveItem()](#OneDrive+uploadDriveItem)
        * [.getWorkbook()](#OneDrive+getWorkbook)
        * [.listSubscriptions()](#OneDrive+listSubscriptions)
        * [.createSubscription()](#OneDrive+createSubscription)
        * [.refreshSubscription()](#OneDrive+refreshSubscription)
        * [.deleteSubscription()](#OneDrive+deleteSubscription)
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
| [opts.clientSecret] | <code>string</code> | The client secret of the app |
| [opts.refreshToken] | <code>string</code> | The refresh token. |
| [opts.accessToken] | <code>string</code> | The access token. |
| [opts.username] | <code>string</code> | Username for username/password authentication. |
| [opts.password] | <code>string</code> | Password for username/password authentication. |
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
**Params**: <code>function</code> [onCode] - optional function that gets invoked after code was retrieved.  
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
<a name="OneDrive+getDriveRootItem"></a>

### oneDrive.getDriveRootItem()
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
<a name="OneDrive+getWorkbook"></a>

### oneDrive.getWorkbook()
**Kind**: instance method of [<code>OneDrive</code>](#OneDrive)  
<a name="OneDrive+listSubscriptions"></a>

### oneDrive.listSubscriptions()
**Kind**: instance method of [<code>OneDrive</code>](#OneDrive)  
<a name="OneDrive+createSubscription"></a>

### oneDrive.createSubscription()
**Kind**: instance method of [<code>OneDrive</code>](#OneDrive)  
<a name="OneDrive+refreshSubscription"></a>

### oneDrive.refreshSubscription()
**Kind**: instance method of [<code>OneDrive</code>](#OneDrive)  
<a name="OneDrive+deleteSubscription"></a>

### oneDrive.deleteSubscription()
**Kind**: instance method of [<code>OneDrive</code>](#OneDrive)  
<a name="OneDrive+fetchChanges"></a>

### oneDrive.fetchChanges(resource, [token]) ⇒ <code>Promise.&lt;Array&gt;</code>
Fetches the changes from the respective resource using the provided delta token.
Use an empty token to fetch the initial state or `latest` to fetch the latest state.

**Kind**: instance method of [<code>OneDrive</code>](#OneDrive)  
**Returns**: <code>Promise.&lt;Array&gt;</code> - An object with an array of changes and a delta token.  

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

<a name="OneDriveMock"></a>

## OneDriveMock
Mock OneDrive client that supports some of the operations the `OneDrive` class does.

**Kind**: global class  

* [OneDriveMock](#OneDriveMock)
    * [.registerWorkbook(driveId, itemId, data)](#OneDriveMock+registerWorkbook) ⇒ [<code>OneDriveMock</code>](#OneDriveMock)
    * [.registerShareLink(uri, driveId, itemId)](#OneDriveMock+registerShareLink) ⇒ [<code>OneDriveMock</code>](#OneDriveMock)
    * [.getDriveItemFromShareLink()](#OneDriveMock+getDriveItemFromShareLink)
    * [.getWorkbook()](#OneDriveMock+getWorkbook)
    * [.getClient()](#OneDriveMock+getClient)

<a name="OneDriveMock+registerWorkbook"></a>

### oneDriveMock.registerWorkbook(driveId, itemId, data) ⇒ [<code>OneDriveMock</code>](#OneDriveMock)
Register a mock workbook.

**Kind**: instance method of [<code>OneDriveMock</code>](#OneDriveMock)  
**Returns**: [<code>OneDriveMock</code>](#OneDriveMock) - this  

| Param | Type | Description |
| --- | --- | --- |
| driveId | <code>string</code> | The drive id |
| itemId | <code>string</code> | the item id |
| data | <code>object</code> | Mock workbook data |

<a name="OneDriveMock+registerShareLink"></a>

### oneDriveMock.registerShareLink(uri, driveId, itemId) ⇒ [<code>OneDriveMock</code>](#OneDriveMock)
Register a mock sharelink.

**Kind**: instance method of [<code>OneDriveMock</code>](#OneDriveMock)  
**Returns**: [<code>OneDriveMock</code>](#OneDriveMock) - this;  

| Param | Type | Description |
| --- | --- | --- |
| uri | <code>string</code> | The sharelink uri |
| driveId | <code>string</code> | The drive id |
| itemId | <code>string</code> | The the item id |

<a name="OneDriveMock+getDriveItemFromShareLink"></a>

### oneDriveMock.getDriveItemFromShareLink()
**Kind**: instance method of [<code>OneDriveMock</code>](#OneDriveMock)  
**See**: OneDrive#getDriveItemFromShareLink  
<a name="OneDriveMock+getWorkbook"></a>

### oneDriveMock.getWorkbook()
**Kind**: instance method of [<code>OneDriveMock</code>](#OneDriveMock)  
**See**: OneDrive#getWorkbook  
<a name="OneDriveMock+getClient"></a>

### oneDriveMock.getClient()
**Kind**: instance method of [<code>OneDriveMock</code>](#OneDriveMock)  
**See**: OneDrive#getClient  
<a name="getActualError"></a>

## getActualError(e)
Returns the actual error, recursively descending through all error properties.

**Kind**: global function  

| Param | Type | Description |
| --- | --- | --- |
| e | <code>Error</code> | error caught |

<a name="handleNamedItems"></a>

## handleNamedItems(sheet, segs, method, body) ⇒ <code>object</code>
Handle the `namedItems` operation on a workbook / worksheet

**Kind**: global function  
**Returns**: <code>object</code> - The response value  

| Param | Type | Description |
| --- | --- | --- |
| sheet | <code>object</code> | The mock data |
| segs | <code>Array.&lt;string&gt;</code> | Array of path segments |
| method | <code>string</code> | Request method |
| body | <code>object</code> | Request body |

<a name="handleTable"></a>

## handleTable(sheet, segs, method, body) ⇒ <code>object</code>
Handle the `table` operation on a workbook / worksheet

**Kind**: global function  
**Returns**: <code>object</code> - The response value  

| Param | Type | Description |
| --- | --- | --- |
| sheet | <code>object</code> | The mock data |
| segs | <code>Array.&lt;string&gt;</code> | Array of path segments |
| method | <code>string</code> | Request method |
| body | <code>object</code> | Request body |

<a name="getActualError"></a>

## getActualError(e)
Returns the actual error, recursively descending through all error properties.

**Kind**: global function  

| Param | Type | Description |
| --- | --- | --- |
| e | <code>Error</code> | error caught |

<a name="getActualError"></a>

## getActualError(e)
Returns the actual error, recursively descending through all error properties.

**Kind**: global function  

| Param | Type | Description |
| --- | --- | --- |
| e | <code>Error</code> | error caught |

<a name="driveItemToURL"></a>

## driveItemToURL(driveItem) ⇒ <code>URL</code>
Returns a onedrive uri for the given drive item. the uri has the format:
`onedrive:/drives/<driveId>/items/<itemId>`

**Kind**: global function  
**Returns**: <code>URL</code> - An url representing the drive item  

| Param | Type |
| --- | --- |
| driveItem | <code>DriveItem</code> | 

<a name="driveItemFromURL"></a>

## driveItemFromURL(url) ⇒ <code>DriveItem</code>
Returns a partial drive item from the given url. The urls needs to have the format:
`onedrive:/drives/<driveId>/items/<itemId>`. if the url does not start with the correct
protocol, {@code null} is returned.

**Kind**: global function  
**Returns**: <code>DriveItem</code> - A (partial) drive item.  

| Param | Type | Description |
| --- | --- | --- |
| url | <code>URL</code> \| <code>string</code> | The url of the drive item. |

