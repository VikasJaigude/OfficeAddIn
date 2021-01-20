// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
/* 
    This file provides the provides functionality to get Microsoft Graph data. 
*/

var myModule = require('./odata-helper');

let domain = "graph.microsoft.com";
let versionURLsegment = "/v1.0";

// If any part of queryParamsSegment comes from user input,
// be sure that it is sanitized so that it cannot be used in
// a Response header injection attack.

async function deleteDraftMessage(accessToken, apiURLsegment, queryParamsSegment) {
    console.log('inside deleteDraftMessage');
    return new Promise(async (resolve, reject) => { 

        try {
            const oData = await myModule.deleteDraft(accessToken, domain, apiURLsegment, versionURLsegment, queryParamsSegment);
            resolve(oData);
        }
        catch(error) {
            reject(Error("Unable to call Microsoft Graph. " + error.toString()));
        }
    })        
} 

async function makeGraphApiCall(accessToken, itemId) {
    return new Promise(async (resolve, reject) => { 
            try {
                const oData = await myModule.callGraphApi(accessToken, "/DeleteDraft", itemId);
                resolve(oData);
            }
            catch(error) {
                reject(Error("Unable to call Microsoft Graph. " + error.toString()));
            }
        })    

    //     let res;
    //     $.ajax({type: "GET", 
    //         url: "/DeleteDraft",
    //         headers: {"access_token": accessToken, "Id": itemId },
    //         cache: false
    //     }).done(function (response) {

    //     console.log(Response);
    //     res = response; 
    //     })
    //     .fail(function (errorResult) {
    //         // This error is relayed from `app.get('/getuserdata` in app.js file.
    //         console.log("Error from Microsoft Graph: " + JSON.stringify(errorResult));
    // });
return res;
}


module.exports = {
    deleteDraftMessage: deleteDraftMessage,
    makeGraphApiCall: makeGraphApiCall
}