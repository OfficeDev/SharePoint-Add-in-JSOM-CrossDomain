// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.


// The allAnnouncements variable is used by more than one 
//  function to retrieve and process the results.
var allAnnouncements;
var hostweburl;
var addinweburl;

// Load the required SharePoint libraries
$(document).ready(function () {
    //Get the URI decoded URLs.
    hostweburl =
        decodeURIComponent(
            getQueryStringParameter("SPHostUrl")
    );
    addinweburl =
        decodeURIComponent(
            // The parameter is SPAppWebUrl, not SPAdd-inWebUrl because add-ins were originally called "apps".
            getQueryStringParameter("SPAppWebUrl")
    );

    // resources are in URLs in the form:
    // web_url/_layouts/15/resource
    var scriptbase = hostweburl + "/_layouts/15/";

    // Load the js files and continue to the successHandler
    $.getScript(scriptbase + "SP.Runtime.js",
        function () {
            $.getScript(scriptbase + "SP.js",
                function () { $.getScript(scriptbase + "SP.RequestExecutor.js", execCrossDomainRequest); }
                );
        }
        );
});

// Function to prepare and issue the request to get
//  SharePoint data
function execCrossDomainRequest() {
    // context: The ClientContext object provides access to
    //      the web and lists objects.
    // factory: Initialize the factory object with the
    //      app web URL.
    var context = new SP.ClientContext(addinweburl);
    var factory =
        new SP.ProxyWebRequestExecutorFactory(
            addinweburl
        );
    context.set_webRequestExecutorFactory(factory);

    //Get the web and list objects
    //  and prepare the query
    var web = context.get_web();
    var list = web.get_lists().getByTitle("Announcements");
    var camlString =
        "<View><ViewFields>" +
            "<FieldRef Name='Title' />" +
            "<FieldRef Name='Body' />" +
        "</ViewFields></View>";

    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(camlString);
    allAnnouncements = list.getItems(camlQuery);

    context.load(allAnnouncements, "Include(Title, Body)");

    //Execute the query with all the previous 
    //  options and parameters
    context.executeQueryAsync(
        successHandler, errorHandler
    );
}

// Function to handle the success event.
// Prints the data to the page.
function successHandler(data, req) {
    var announcementsHTML = "";
    var enumerator = allAnnouncements.getEnumerator();

    while (enumerator.moveNext()) {
        var announcement = enumerator.get_current();
        announcementsHTML = announcementsHTML +
            "<p><h1>" + announcement.get_item("Title") +
            "</h1>" + announcement.get_item("Body") +
            "</p><hr>";
    }

    document.getElementById("renderAnnouncements").innerHTML =
        announcementsHTML;
}

// Function to handle the error event.
// Prints the error message to the page.
function errorHandler(data, error, errorMessage) {
    document.getElementById("renderAnnouncements").innerText =
        "Could not complete cross-domain call: " +
        errorMessage;
}

// Function to retrieve a query string value.
// For production purposes you may want to use
//  a library to handle the query string.
function getQueryStringParameter(paramToRetrieve) {
    var params =
        document.URL.split("?")[1].split("&");
    var strParams = "";
    for (var i = 0; i < params.length; i = i + 1) {
        var singleParam = params[i].split("=");
        if (singleParam[0] == paramToRetrieve)
            return singleParam[1];
    }
}

/*

SharePoint-Add-in-JSOM-CrossDomain, https://github.com/OfficeDev/SharePoint-Add-in-JSOM-CrossDomain
 
Copyright (c) Microsoft Corporation
All rights reserved. 
 
MIT License:
Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:
 
The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.
 
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.    
  
*/