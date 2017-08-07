# sp-rest-api.js

## A lightweight JS library to work with SharePoint lists using the SharePoint REST API.

The purpose of `sp-rest-api.js` is to simplify queries to SharePoint REST API by providing a few shortcuts and convenient defaults. 

### Examples

#### Get all items from a list

```
// Read all items from the Projects list in SharePoint,
// and print them using console.log(), the default callback.
var api = new SpRestApi();
api.lists('Projects').getAllItems();
```

#### Get a single item from a list, run a callback

```
// Read a single SharePoint list item, and run a callback
// after success or error.
var api = new SpRestApi();
var options = {
    listTitle: 'Projects',
    onsuccess: function(data) {
        alert('Success!');
    },
    onerror: function(data) {
        alert('Error!');
    },
};
api.options(options).getItem(123);
```


### Requirements

- jQuery 1+
- SharePoint 2013, 2016 or online

### Setup

Just place `sp-rest-api.js` into any folder on the site, e.g. into `/SiteAssets`, and include it after the jQuery and SP JavaScript files:

```
<script src='../SiteAssets/sp-rest-api.js'></script>
```

### License

MIT. Do whatever you want.
