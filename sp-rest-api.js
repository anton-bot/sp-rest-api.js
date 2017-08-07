/**
 * @fileoverview A libary for working with SharePoint lists via the SharePoint
 * REST API. Supports reading, updating and deleting list items, as well as a
 * few extra functions such as getting a list of all columns in a SP list.
 */

/**
 * @class SpRestApi
 * @typedef {Object} SpRestApi
 * Contains methods for calling the SharePoint REST API.
 */
var SpRestApi = function () {

};

/**
 * @typedef {Object} SpRestApiOptions - The options passed to methods
 *      inside the SpRestApi class.
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
 * @property {string} [verbosity] - The amount of metadata to be returned
 *      in the JSON response from server. Use the SpRestApi.Verbosity enum.
 * @property {string} [siteUrl] - The SharePoint site URL which is usually
 *      obtained from the _spPageContextInfo.webAbsoluteUrl.
 * @property {Array.<string>} [urls] - The URLs of various API calls, e.g. to
 *      get a list item, all items in a list etc.
 */

/**
 * The default options that will be used unless overridden.
 * @type {SpRestApiOptions}
 */
SpRestApi.prototype.defaultOptions = {
    onsuccess: console.log,
    onerror: console.log,
    listTitle: '',
    maxItems: 5000, 
    recursiveFetch: true,
    verbosity: SpRestApi.Verbosity.COMPACT,
    siteUrl: _spPageContextInfo.webAbsoluteUrl,
    urls: {
        list: '/_api/web/lists/getbytitle(\'{0}\')/items',
        item: '/_api/web/lists/getbytitle(\'{0}\')/items({1})',
    },
};

/**
 * Determines the amount of metadata in the response JSON from server.
 * For most cases, 'COMPACT' should be enough.
 * @readonly
 * @enum {string}
 * @typedef Verbosity
 */
SpRestApi.prototype.Verbosity = {
    /** Values are placed in `.value`. Less metadata. */
    COMPACT: 'application/json; odata=nometadata',

    /** Values are placed in `.d`. More metadata. */
    VERBOSE: 'application/json; odata=verbose',
};

/**
 * Sets the list title (list display name) of this SpRestApi instance.
 * The .lists() must be called before any call to other list-related
 * methods.
 * @param {string} listTitle - The display name of the SharePoint list.
 * @returns {SpRestApi} Returns the instance of this SpRestApi object.
 */
SpRestApi.prototype.lists = function (listTitle) {
    this.listTitle = listTitle;
    return this;
};

/**
 * Sets the SpRestApiOptions. If not called, before the request to server,
 * the default options will be used.
 * @param {SpRestApiOptions} options - The partial SpRestApiOptions object,
 *      where each field will override a default setting in `.defaultOptions`.
 * @returns {SpRestApi} This SpRestApi instance.
 */
SpRestApi.prototype.options = function (options) {
    // Merge the specified options with the default options
    this.options = $.extend(this.defaultOptions, options);
    return this;
};

/**
 * Returns all items from a list, or all items up to the SharePoint limit
 * or the limit specified in the options.
 */
SpRestApi.prototype.getAllItems = function () {
    var url = this.options.urls.list.format(this.options.listTitle);
    this.loadUrl(url, 'GET', options.onsuccess, options.onerror);
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
 */
SpRestApi.prototype.loadUrl = function (url, method, success, error) {
    $.ajax({
        url: url,
        type: method,
        cache: false,
        contentType: this.options.verbosity,
        headers: {
            "Accept": this.options.verbosity,
            "X-RequestDigest":
                document.getElementById("__REQUESTDIGEST").value // TODO FIXME get via a separate function or property
        },
        success: function (data) {
            success(data);
        },
        error: function (data) {
            error(data);
        },
    });
};