## Classes

<dl>
<dt><a href="#OneDrive">OneDrive</a></dt>
<dd><p>Helper class that facilitates accessing one drive.</p>
</dd>
<dt><a href="#OneDriveAuth">OneDriveAuth</a></dt>
<dd><p>Helper class that facilitates accessing one drive.</p>
</dd>
<dt><a href="#OneDriveMock">OneDriveMock</a></dt>
<dd><p>Mock OneDrive client that supports some of the operations the <code>OneDrive</code> class does.</p>
</dd>
<dt><a href="#SharePointSite">SharePointSite</a></dt>
<dd><p>Helper class accessing folders and files using the SharePoint V1 API.</p>
</dd>
</dl>

## Constants

<dl>
<dt><a href="#globalTenantCache">globalTenantCache</a> : <code>Map.&lt;string, string&gt;</code></dt>
<dd><p>map that caches the tenant ids</p>
</dd>
</dl>

## Functions

<dl>
<dt><a href="#handleNamedItems">handleNamedItems(sheet, segs, method, body)</a> ⇒ <code>object</code></dt>
<dd><p>Handle the <code>namedItems</code> operation on a workbook / worksheet</p>
</dd>
<dt><a href="#handleTable">handleTable(sheet, segs, method, body)</a> ⇒ <code>object</code></dt>
<dd><p>Handle the <code>table</code> operation on a workbook / worksheet</p>
</dd>
<dt><a href="#splitByExtension">splitByExtension(name)</a> ⇒ <code>Array.&lt;string&gt;</code></dt>
<dd><p>Splits the given name at the last &#39;.&#39;, returning the extension and the base name.</p>
</dd>
<dt><a href="#sanitizeName">sanitizeName(name)</a> ⇒ <code>string</code></dt>
<dd><p>Sanitizes the given string by :</p>
<ul>
<li>convert to lower case</li>
<li>normalize all unicode characters</li>
<li>replace all non-alphanumeric characters with a dash</li>
<li>remove all consecutive dashes</li>
<li>remove all leading and trailing dashes</li>
</ul>
</dd>
<dt><a href="#sanitizePath">sanitizePath(filepath, opts)</a> ⇒ <code>string</code></dt>
<dd><p>Sanitizes the file path by:</p>
<ul>
<li>convert to lower case</li>
<li>normalize all unicode characters</li>
<li>replace all non-alphanumeric characters with a dash</li>
<li>remove all consecutive dashes</li>
<li>remove all leading and trailing dashes</li>
</ul>
<p>Note that only the basename of the file path is sanitized. i.e. The ancestor path and the
extension is not affected.</p>
</dd>
<dt><a href="#editDistance">editDistance(s0, s1)</a> ⇒ <code>number</code> | <code>*</code></dt>
<dd><p>Compute the edit distance using a recursive algorithm. since we only expect to have relative
short filenames, the algorithm shouldn&#39;t be too expensive.</p>
</dd>
<dt><a href="#superTrim">superTrim(str)</a> ⇒ <code>string</code></dt>
<dd><p>Trims the string at both ends and removes the zero width unicode chars:</p>
<ul>
<li>U+200B zero width space</li>
<li>U+200C zero width non-joiner Unicode code point</li>
<li>U+200D zero width joiner Unicode code point</li>
<li>U+FEFF zero width no-break space Unicode code point</li>
</ul>
</dd>
</dl>

## Typedefs

<dl>
<dt><a href="#AuthenticationResult">AuthenticationResult</a> : <code>module:@azure/msal-node~AuthenticationResult</code></dt>
<dd><p>aliases</p>
</dd>
</dl>

<a name="globalTenantCache"></a>

## globalTenantCache : <code>Map.&lt;string, string&gt;</code>
map that caches the tenant ids

**Kind**: global constant  
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

<a name="splitByExtension"></a>

## splitByExtension(name) ⇒ <code>Array.&lt;string&gt;</code>
Splits the given name at the last '.', returning the extension and the base name.

**Kind**: global function  
**Returns**: <code>Array.&lt;string&gt;</code> - Returns an array containing the base name and extension.  

| Param | Type | Description |
| --- | --- | --- |
| name | <code>string</code> | Filename |

<a name="sanitizeName"></a>

## sanitizeName(name) ⇒ <code>string</code>
Sanitizes the given string by :
- convert to lower case
- normalize all unicode characters
- replace all non-alphanumeric characters with a dash
- remove all consecutive dashes
- remove all leading and trailing dashes

**Kind**: global function  
**Returns**: <code>string</code> - sanitized name  

| Param | Type |
| --- | --- |
| name | <code>string</code> | 

<a name="sanitizePath"></a>

## sanitizePath(filepath, opts) ⇒ <code>string</code>
Sanitizes the file path by:
- convert to lower case
- normalize all unicode characters
- replace all non-alphanumeric characters with a dash
- remove all consecutive dashes
- remove all leading and trailing dashes

Note that only the basename of the file path is sanitized. i.e. The ancestor path and the
extension is not affected.

**Kind**: global function  
**Returns**: <code>string</code> - sanitized file path  

| Param | Type | Description |
| --- | --- | --- |
| filepath | <code>string</code> | the file path |
| opts | <code>object</code> | Options |
| [opts.ignoreExtension] | <code>boolean</code> | if {@code true} ignores the extension |

<a name="editDistance"></a>

## editDistance(s0, s1) ⇒ <code>number</code> \| <code>\*</code>
Compute the edit distance using a recursive algorithm. since we only expect to have relative
short filenames, the algorithm shouldn't be too expensive.

**Kind**: global function  

| Param | Type | Description |
| --- | --- | --- |
| s0 | <code>string</code> | Input string |
| s1 | <code>string</code> | Input string |

<a name="superTrim"></a>

## superTrim(str) ⇒ <code>string</code>
Trims the string at both ends and removes the zero width unicode chars:

- U+200B zero width space
- U+200C zero width non-joiner Unicode code point
- U+200D zero width joiner Unicode code point
- U+FEFF zero width no-break space Unicode code point

**Kind**: global function  
**Returns**: <code>string</code> - trimmed and stripped string  

| Param | Type | Description |
| --- | --- | --- |
| str | <code>string</code> | input string |

<a name="AuthenticationResult"></a>

## AuthenticationResult : <code>module:@azure/msal-node~AuthenticationResult</code>
aliases

**Kind**: global typedef  
