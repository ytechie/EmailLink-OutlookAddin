/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    var linkDisplay = document.getElementById("linkDisplay");
    linkDisplay.innerHTML = "Loading...";

    var item = Office.context.mailbox.item;

    var link = 'https://outlook.office.com/owa/?ItemID=';
    link += item.itemId.toString();
    link += "&viewmodel=ReadMessageItem&path=&exvsurl=1";
  
    linkDisplay.innerText = link;
  }
});
