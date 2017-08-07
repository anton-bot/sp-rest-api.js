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
    /**
     * @typedef {Object} SpRestApiOptions - The options passed to methods
     *      inside the SpRestApi class.
     * @property {Function} onsuccess - The callback function for successful
     *      requests to SharePoint REST API.
     * @property {Function} onerror - The callback function for failed requests
     *      to SharePoint REST API.
     * @property {number} maxItems - The maximum number of items to be returned
     *      from the list. If `recursiveFetch` is set to true, this is the
     *      maximum number of items to fetch on each request to server. If not
     *      not specified, defaults to SharePoint's limit of 100. Maximum is
     *      5000 due to SharePoint limitations.
     * @property {boolean} recursiveFetch - Fetch all items from the list by
     *      repeatedly making server requests until all list items are fetched.
     *      This is to overcome SharePoint's limitation of maximum 5000 items
     *      per call.
     */
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
}

/**
 * Sets the SpRestApiOptions. If not called, before the request to server,
 * the default options will be used.
 * @param {SpRestApiOptions} options
 */
SpRestApi.prototype.options = function (options) {
    // TODO FIXME merge with default options
    this.options = options;
};

/**
 * Returns all items from a list, or all items up to the SharePoint limit
 * or the limit specified in the options.
 */
SpRestApi.prototype.getAllItems = function () {

};

/**
 * A generic function to call any URL of the SharePoint REST API. Usually
 * there is no need to call this method directly. 
 */
SpRestApi.prototype.loadUrl: function (url, method, success, error) {
    $.ajax({
        url: url,
        type: method,
        cache: false,
        contentType: "application/json; odata=verbose",
        headers: {
            "Accept": "application/json; odata=verbose",
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