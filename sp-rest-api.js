/**
 * @fileoverview A libary for working with SharePoint lists via the SharePoint
 * REST API. Supports reading, updating and deleting list items, as well as a
 * few extra functions such as getting a list of all columns in a SP list.
 */

/**
 * Initializes a new instance of the SpRestApi class, which contains methods
 * for calling the SharePoint REST API.
 * @class SpRestApi
 * @typedef {Object} SpRestApi
 * @constructor
 * @param {SpRestApiOptions} [options] - The optional settings to override the
 *      defaults.
 */
var SpRestApi = function (options) {
    /**
     * The default options that will be used unless overridden.
     * @type {SpRestApiOptions}
     */
    this.defaultOptions = {
        onsuccess: console.log,
        onerror: console.log,
        listTitle: '',
        maxItems: 100,
        recursiveFetch: true,
        verbosity: SpRestApi.Verbosity.VERBOSE,
        siteUrl: _spPageContextInfo ? _spPageContextInfo.webAbsoluteUrl : '',
        token: document.getElementById("__REQUESTDIGEST") ? 
            document.getElementById("__REQUESTDIGEST").value : '',
        urls: {
            list: '/_api/web/lists/getbytitle(\'{0}\')/items',
            item: '/_api/web/lists/getbytitle(\'{0}\')/items({1})',
            user: '/_api/Web/GetUserById({0})?$expand=Groups',
        },
    };

    this.options = $.extend(this.defaultOptions, options);
};

/**
 * Determines the amount of metadata in the response JSON from server.
 * For most cases, 'COMPACT' should be enough.
 * @readonly
 * @enum {string}
 * @typedef Verbosity
 */
SpRestApi.Verbosity = {
    /** Values are placed in `.value`. Less metadata. 
        Not supported in SP 2013 and earlier. */
    COMPACT: 'application/json;odata=nometadata',

    /** Value are placed in `.value`. Moderate metadata. 
        Not supported in SP 2013 and earlier. */
    MINIMAL: 'application/json;odata=minimalmetadata',

    /** Values are placed in `.d.results`. More metadata. */
    VERBOSE: 'application/json;odata=verbose',
};

/**
 * @typedef {Object} SpRestApiOptions - The options passed to methods
 *      inside the SpRestApi class.
 * @property {Array.<SpRestApiFilterCriterion>} filters - The filters to be
 *      used in filtering the list items using the $filter parameter.
 * @property {string} [token] - The SharePoint's request digest, a token that
 *      it is required for every API call. Required only if using outside a
 *      SharePoint page, otherwise it is obtained automatically from DOM.
 * @property {Function} [onsuccess] - The callback function for successful
 *      requests to SharePoint REST API.
 * @property {Function} [onerror] - The callback function for failed requests
 *      to SharePoint REST API.
 * @property {string} [listTitle] - The display name of the SharePoint list.
 * @property {number} [maxItems] - The maximum number of items to be returned
 *      from the list. If `recursiveFetch` is set to true, this is the
 *      maximum number of items to fetch on each request to server. If not
 *      not specified, defaults to SharePoint's limit of 100. Maximum is
 *      5000 due to SharePoint limitations.
 * @property {boolean} [recursiveFetch] - Fetch all items from the list by
 *      repeatedly making server requests until all list items are fetched.
 *      This is to overcome SharePoint's limitation of maximum 5000 items
 *      per call.
 * @property {Verbosity} [verbosity] - The amount of metadata to be returned
 *      in the JSON response from server. Use the SpRestApi.Verbosity enum.
 * @property {string} [siteUrl] - The SharePoint site URL which is usually
 *      obtained from the _spPageContextInfo.webAbsoluteUrl. Required if using
 *      this library outside of a SharePoint page.
 * @property {Array.<string>} [urls] - The URLs of various API calls, e.g. to
 *      get a list item, all items in a list etc.
 */
/**
 * Stores this instance's options.
 * @type {SpRestApiOptions}
 */
SpRestApi.prototype.options = {};

/**
 * Array of cached SharePoint list items. Stores the data during the recursive
 * fetch (options.recursiveFetch = true), and is cleared after another
 * fetch operation begins.
 * @type {Array.<Object>}
 */
SpRestApi.prototype.cachedListItems = [];

/**
 * Sets the list title (list display name) of this SpRestApi instance.
 * The .lists() must be called before any call to other list-related
 * methods. Equivalent to .config({listTitle: 'List Name'}).
 * @param {string} listTitle - The display name of the SharePoint list.
 * @returns {SpRestApi} Returns the instance of this SpRestApi object.
 */
SpRestApi.prototype.lists = function (listTitle) {
    this.options.listTitle = listTitle;
    return this;
};

/**
 * Sets the SpRestApiOptions. If not called, before the request to server,
 * the default options will be used.
 * @param {SpRestApiOptions} options - The partial SpRestApiOptions object,
 *      where each field will override a default setting in `.defaultOptions`.
 * @returns {SpRestApi} This SpRestApi instance.
 */
SpRestApi.prototype.config = function (options) {
    // Merge the specified options with the default options
    this.options = $.extend(this.defaultOptions, options);
    return this;
};

/**
 * Adds the $top= string to the specified URL, to limit the max number of
 * items in the response.
 * @param {string} url - the URL to add the $top string to.
 * @returns {string} The URL with the added "?"/"&" character and the $top=
 *      URL parameter.
 */
SpRestApi.prototype.addMaxItems = function (url) {
    // Add '?' or '&' to URL query string
    url += url.includes('?') ? '&' : '?';

    return url + "$top=" + this.options.maxItems;
};

/**
 * Returns all items from a list, or all items up to the SharePoint limit
 * or the limit specified in the options.
 */
SpRestApi.prototype.getAllItems = function () {
    var url = this.generateGetAllListItemsUrl();

    this.cachedListItems = []; // reset cached items for recursive fetching

    var onsuccess;
    if (this.options.recursiveFetch) {
        // Continue fetching items recursively until we load the entire list:
        onsuccess = $.proxy(this.continueRecursiveFetch, this);
    } else {
        onsuccess = this.options.onsuccess;        
    }

    this.loadUrl(url, 'GET', onsuccess, this.options.onerror);
};

/**
 * Generates the URL for fetching all items from a list. Uses in getAllItems()
 * and in getAllItemsFromSubfolder(), as well as to create/POST a new item.
 * @returns {string} The SP API URL for fetching all items from a list.
 */
SpRestApi.prototype.generateGetAllListItemsUrl = function () {
    var url = this.options.siteUrl +
        this.options.urls.list.format(this.options.listTitle);
    url = this.addMaxItems(url);
    return url;
};

/**
 * Generates the URL for getting, updating or deleting a single list item.
 * @param {number} itemId - The list item ID of the item to be deleted.
 * @returns {string} An URL ending with e.g. /Lists('MyList')/Items(761)
 */
SpRestApi.prototype.generateSingleListItemUrl = function (itemId) {
    return this.options.siteUrl +
        this.options.urls.item.format(this.options.listTitle, itemId);
};

/**
 * Returns all list items from a subfolder of a SharePoint list. Uses a hacky
 * way - by matching a substring of the FileRef, to avoid relying on CAML.
 * @param {string} subfolderName - The display name of the subfolder in a list.
 *      Must not include any slashes.
 */
SpRestApi.prototype.getAllItemsFromListSubfolder = function (subfolderName) {
    // The idea here is that if the FileRef property of a list item is
    // like this - /sites/mysite/Lists/My List/My Subfolder/123_.000 - 
    // then we can filter the list by FileRef containing the 
    // string 'Lists/My List/My Subfolder'.
    var fileref = 'Lists/' +
        this.options.listTitle + '/' + subfolderName + '/';

    var url = this.generateGetAllListItemsUrl();
    url += '&$filter=substringof(\'' + fileref + '\', FileRef)';

    this.cachedListItems = []; // reset cached items for recursive fetching

    var onsuccess;
    if (this.options.recursiveFetch) {
        // Continue fetching items recursively until we load the entire list:
        onsuccess = $.proxy(this.continueRecursiveFetch, this); // save `this`
    } else {
        onsuccess = this.options.onsuccess;
    }

    this.loadUrl(url, 'GET', onsuccess, this.options.onerror);
};

/**
 * Keeps loading data recursively until all list items are obtained. This is
 * to overcome SharePoint's limitation on the number of list items per query.
 */
SpRestApi.prototype.continueRecursiveFetch = function (data) {
    var newData;
    var nextUrl;

    // Depending on the Verbosity setting, the data will be returned in
    // significantly different formats (inside data.d.results or data.value).
    if (this.options.verbosity === SpRestApi.Verbosity.VERBOSE) {
        newData = data.d.results;
        nextUrl = data.d.__next;
    } else {
        newData = data.value;
        nextUrl = data['odata.nextLink']; // dot in property name
    }

    this.cachedListItems = this.cachedListItems.concat(newData);

    if (nextUrl) {
        // While next URL is not empty, keep loading recursively. 
        // Preserve `this` inside continueRecursiveFetch using $.proxy.
        this.loadUrl(nextUrl, 'GET',
            $.proxy(this.continueRecursiveFetch, this), this.options.onerror);
    } else {
        // Load complete - generate a structure similar to server response, 
        // and call the callback.
        var entireData;
        if (this.options.verbosity === SpRestApi.Verbosity.VERBOSE) {
            entireData = {
                d: {
                    results: this.cachedListItems
                }
            };
        } else {
            entireData = {
                value: this.cachedListItems
            };
        }

        this.options.onsuccess(entireData);
    }
};

/**
 * Returns a single item from a list.
 * @param {number} itemId - The SharePoint list item ID of the item we need to
 *      fetch.
 */
SpRestApi.prototype.getItem = function (itemId) {
    if (!itemId) { throw "The list item ID must not be empty."; }

    var url = this.options.siteUrl +
        this.options.urls.item.format(this.options.listTitle, itemId);
    this.loadUrl(url, 'GET', this.options.onsuccess, this.options.onerror);
};

/**
 * Creates a new item in a SharePoint list. 
 * @param {Object} item - The new SharePoint list item to be created. No need
 *      to add the __metadata attribute.
 */
SpRestApi.prototype.createItem = function (item) {
    item.__metadata = {
        "type": SpRestApi.getListItemType(this.options.listTitle)
    };

    var url = this.generateGetAllListItemsUrl();

    this.loadUrl(url, 'POST',
        this.options.onsuccess, this.options.onerror, item);
};

/**
 * Updates a single item in the SharePoint list. Only the values in columns
 * specified in `data` will be overwritten.
 * @param {number} listItemId - The SharePoint list item ID of the item to be
 *      updated.
 * @param {Object} item - The partial or full SharePoint list item containing
 *      only the columns that need to be replaced.
 */
SpRestApi.prototype.updateItem = function (listItemId, item) {
    item.__metadata = {
        "type": SpRestApi.getListItemType(this.options.listTitle)
    };

    var url = this.generateSingleListItemUrl(listItemId);

    this.loadUrl(url, 'MERGE',
        this.options.onsuccess, this.options.onerror, item);
};

/**
 * Deletes a single item from a SharePoint list.
 * @param {number} itemId - The SharePoint list item ID of the item to be
 *      deleted.
 */
SpRestApi.prototype.deleteItem = function (itemId) {
    var url = this.generateSingleListItemUrl(itemId);
    this.loadUrl(url, 'DELETE', this.options.onsuccess, this.options.onerror);
};

/**
 * Fetches the SharePoint user information, including the Groups that he is a
 * member of.
 * @param {number} userId - The SharePoint user ID of the user whose information
 *      we are requesting.
 */
SpRestApi.prototype.getUserById = function (userId) {
    if (!userId) {
        throw "Tried to get user information using an empty user ID.";
    }

    var url = this.options.urls.user.format(userId);

    this.loadUrl(url, 'GET', this.options.onsuccess, this.options.onerror);
};

/**
 * Fetches the information about the current user, such as email, groups etc.
 * Wrapper for getUserById(), gets the current user ID automatically.
 */
SpRestApi.prototype.getCurrentUser = function () {
    if (!_spPageContextInfo) { throw 'Missing _spPageContextInfo.'; }

    this.getUserById(_spPageContextInfo.userId);
};

/**
 * Generates the ListItemType which is required by SharePoint when creating a
 * new list item. It is based on the list name where characters such as spaces
 * are replaced with sequences like _x0020_.
 * @param {string} listTitle - The display name of the list where we are
 *      creating the new list item.
 * @returns {string} A SharePoint list item type which looks like
 *      "SP.Data.ProjectsListItem".
 * @static
 */
SpRestApi.getListItemType = function (listTitle) {
    var type = "SP.Data." + listTitle.capitalize() + "ListItem";
    return SpRestApi.replaceSharepointSpecialChars(type);
};

/**
 * Replaces special characters (like underscores and spaces) that
 * cannot be used in SharePoint list names because SharePoint changes them to 
 * escape sequences, as seen in the function.
 * @param {string} inputString - A string that may contain an underscore or a
 *      space, which will be replaced.
 * @returns {string} The string that is safe to use in SharePoint list item
 *      internal type.
 * @static
 */
SpRestApi.replaceSharepointSpecialChars = function (inputString) {
    return inputString
        .replace(/_/g, '_x005f_')
        .replace(/ /g, '_x0020_')
        .replace(/&/g, '_x0026_');
};

/**
 * A generic function to call any URL of the SharePoint REST API. Usually
 * there is no need to call this method directly.
 * @param {string} url - The URL of the SharePoint REST API to be queried. Will
 *      not be modified by this method.
 * @param {string} method - HTTP method for this request, e.g. 'GET', 'POST',
 *      'DELETE'.
 * @param {Function} success - Callback for successfull REST API call.
 * @param {Function} error - Callback for failed REST API call.
 * @param {Object} [data] - The data to be POSTed/PUT to the server.
 */
SpRestApi.prototype.loadUrl = function (url, method, success, error, data) {
    // Preserve `this` containing the current instance of SpRestApi.
    var $ajax = $.proxy($.ajax, this);

    // These headers are common for all requests
    var headers = {
        "Accept": this.options.verbosity,
        "X-RequestDigest": this.options.token,
    };

    // For the DELETE/MERGE methods, we actually send the request as POST
    if (method === 'DELETE' || method === 'MERGE') {
        headers['IF-MATCH'] = '*';
        headers['X-HTTP-METHOD'] = method;
        method = 'POST';
    }

    $ajax({
        url: url,
        type: method,
        cache: false,
        data: data ? JSON.stringify(data) : '',
        contentType: this.options.verbosity,
        headers: headers,
        success: function (response) {
            success(response);
        },
        error: function (response) {
            error(response);
        },
    });
};

/**
 * Sends a request to SharePoint to get the SharePoint authorization token,
 * which is required for all data requests. This is necessary in two cases:
 * - We are loading this script from a non-SharePoint page (e.g. normal HTML);
 * - We are refreshing the token (which expires after 30 minutes).
 * @param {Function} callback - The callback to run after the authorization
 *      token is received successfully.
 */
SpRestApi.prototype.refreshDigest = function (callback) {
    var url = this.siteUrl + '/_api/contextinfo';

    // If request succeeded:
    var success = function (data) {
        if (data && data.d && data.d.GetContextWebInformation &&
            data.d.GetContextWebInformation.FormDigestValue) {
            // If Verbosity = odata=verbose:
            this.options = data.d.GetContextWebInformation.FormDigestValue;

        } else if (data && data.value && data.value.GetContextWebInformation &&
            data.value.GetContextWebInformation.FormDigestValue) { // TODO FIXME need to verify this path is correct in SP online
            // If Verbosity = odata=nometadata, odata=minimalmetadata:
            this.options = data.value.GetContextWebInformation.FormDigestValue;
        } else {
            throw 'Unable to obtain SharePoint authorization token.';
        }

        if (typeof callback === 'function') {
            callback();
        }
    };

    // If request failed:
    var error = function () {
        throw 'Unable to obtain SharePoint authorization token';
    };

    this.loadUrl(url, 'POST', success, error);
};


/* Polyfills
-----------------------------*/

if (!String.prototype.format) {
    /**
    * Replaces symbols like {0}, {1} in a string with the values from
    * the arguments.
    * @returns {string} The string with the replaced placeholders.
    */
    String.prototype.format = function () {
        var args = arguments;
        return this.replace(/{(\d+)}/g, function (match, number) {
            return typeof args[number] !== 'undefined'
                ? args[number]
                : match
                ;
        });
    };
}

if (!String.prototype.includes) {
    String.prototype.includes = function (search, start) {
        'use strict';
        if (typeof start !== 'number') {
            start = 0;
        }

        if (start + search.length > this.length) {
            return false;
        } else {
            return this.indexOf(search, start) !== -1;
        }
    };
}

if (!String.prototype.capitalize) {
    /**
     * Capitalizes the first letter of each word in the string.
     */
    String.prototype.capitalize = function () {
        return this.replace(/\b[a-z]/g, function (letter) {
            return letter.toUpperCase();
        });
    };
}