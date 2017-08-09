# sp-rest-api.js

## A lightweight JS library to work with SharePoint lists using the SharePoint REST API.

The purpose of `sp-rest-api.js` is to make it easier to query SharePoint REST API by providing a few shortcuts and convenient defaults for common CRUD operations, such as reading, updating and deleting list items.

There's no need to learn - begin right away with examples below.

### Examples

All you need to get started is:

```html
<script src="sp-rest-api.js"></script>
```

#### Get all items from a list

```js
// Read all items from the Projects list in SharePoint,
// and print them using console.log(), the default callback.
var api = new SpRestApi();
api.lists('Projects').getAllItems();
```

#### Get a single item from a list, run a callback

```js
// Read a single SharePoint list item, and run a callback
// after success or error.
var api = new SpRestApi({
    listTitle: 'Projects',
    onsuccess: function(data) {
        alert('Success!');
    },
    onerror: function(data) {
        alert('Error!');
    },
});
api.getItem(123);
```


### Requirements

- jQuery 1+
- SharePoint 2013, 2016 or online

If using this library in a non-SharePoint page (e.g. a normal HTML file), you need to specify the `siteUrl` option when initializing the `SpRestApi`, and then run `refreshDigest()` to get the authorization token. Until `refreshDigest()` completes successfully, all SharePoint API calls will fail.

```js
var api = new SpRestApi({
    siteUrl: 'http://sharepoint.example.com/sites/mysite',
});
api.refreshDigest(initializePage); // Insert your callback function here
```

### Setup

Just place `sp-rest-api.js` into any folder on the site, e.g. into `/SiteAssets`, and include it after the jQuery and SP JavaScript files:

```html
<script src='../SiteAssets/sp-rest-api.js'></script>
```

If using inside a SharePoint page, the `<script>` tag cannot be placed after SharePoint's `<input id="__REQUESTDIGEST" type="hidden">` tag, otherwise you will need to run `refreshDigest()` manually to get the SP authorization token.

### Full reference

This library is now in development. Target completion date: end of August 2017.

See the [jsdoc](https://github.com/J3QQ4/sp-rest-api.js/blob/master/jsdoc/SpRestApi.html) for full description of methods and options.

#### Initialization
- `new SpRestApi()` - create new instance, set options
- `config()` - set options after the object was created
- `lists()` - sets the list name only (can also be set via `config()`)

#### Working with SharePoint Lists
##### Reading
- `getAllItems()` - fetch all items from a list
- `getAllItemsFromListSubfolder()` - fetch all items from a subfolder in a list
- `getItem()` - fetch a single item from the list

##### Writing
- `deleteItem()` - deletes a single list item.

#### Utilities

- `refreshDigest()` - gets a new the SharePoint security validation / token, and stores it in the `options`.

#### Other/internal methods

- `addMaxItems()` - adds `$top` parameter to URL.
- `loadUrl()` - fetches the specified URL.
- `generateGetAllListItemsUrl()` - generates the API URL to fetch all items from a list
- `continueRecursiveFetch()` - continues fetching all list items if `options.recursiveFetch` is on.

### License and contributing

Public domain. Do whatever you want.

Pull requests, bug reports and feature requests are welcome.
