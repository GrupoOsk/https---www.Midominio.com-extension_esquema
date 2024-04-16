/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, g */

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    // document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    $(function () {
      $(document).on('click', '.panel_body_identificacion_opcion1', function (event) {
        event.preventDefault();
        //IDentificar y cerrar
        var messageObject = {
          messageType: "dialogClosed"
        };
        var jsonMessage = JSON.stringify(messageObject);
        Office.context.ui.messageParent(jsonMessage);
      });
    });
  }
  ;
});