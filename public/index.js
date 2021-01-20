/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

if (!window.Promise) {
    window.Promise = Office.Promise;
  }

  
    var mailboxItem;
    let bootstrapToken; 
    let exchangeResponse ;
    let customHeader;
    let myFirstPromise;
    //Office.onReady(info => { 
    Office.initialize = function (reason) {
        mailboxItem = Office.context.mailbox.item;
        $(document).ready(function () {
            getBootstrapToken();
        });
    };
    // Office.initialize = function (reason) {
        
    // }

    async function getBootstrapToken() {
        try {
            
            bootstrapToken = await OfficeRuntime.auth.getAccessToken({ forMSGraphAccess: true });
            console.log('getBootstrapToken ' + bootstrapToken);
            // The /api/DoSomething controller will make the token exchange and use the
            // access token it gets back to make the call to MS Graph.
        }
        catch (exception) {
        console.log(exception);
        }
    }


    // Entry point for Contoso Message Body Checker add-in before send is allowed.
    // <param name="event">MessageSend event is automatically passed by BlockOnSend code to the function specified in the manifest.</param>
    function validateBody(event) {
        mailboxItem.body.getAsync("html", { asyncContext: event }, checkBodyOnlyOnSendCallBack);
    }

    // Check if the body contains a specific set of blocked words. If it contains the blocked words, block email from being sent. Otherwise allows sending.
    // <param name="asyncResult">MessageSend event passed from the calling function.</param>
    function checkBodyOnlyOnSendCallBack(onSendAsyncResult) {
        Office.context.mailbox.item.internetHeaders.getAsync(
            ["x-custom-addin"],
            function(asyncResult) {
              if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  console.log("Selected headers: " + JSON.stringify(asyncResult.value));
                  
                  if (JSON.stringify(asyncResult.value) === "{}")
                  {
                    console.log("Error getting selected headers: " + JSON.stringify(asyncResult.error));
                    onSendAsyncResult.asyncContext.completed({ allowEvent: true });
                  }
                  else {
                        //bootstrapToken1= asyncResult.value.match(/x-custom-addin:.*/gim)[0].slice(16);
                        mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'asas' });
                            Office.context.mailbox.item.saveAsync(onSendAsyncResult.asyncContext, function(result) 
                            {
                                if (result.status === Office.AsyncResultStatus.Succeeded) {
                                    let response = $.ajax({type: "GET", 
                                                url: "/auth",
                                                connext: onSendAsyncResult.asyncContext,
                                                headers: {"Authorization": "Bearer " + bootstrapToken, "Id": result.value }, 
                                                cache: false
                                            });
                                            Office.context.mailbox.item.close();

                                } else {
                                    console.error(`saveAsync failed with message ${result.error.message}`);
                                }
                            });
                                
                        }
                } 
                else {
                    console.log("Error getting selected headers: " + JSON.stringify(asyncResult.error));
                    asyncResult.asyncContext.completed({ allowEvent: true });
                }
            }
          );
    }   
    