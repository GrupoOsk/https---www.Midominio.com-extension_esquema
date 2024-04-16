/* global global, Office, self, window */

/// <reference path="../jquery.min.js" />
const { send } = require("process");


var intervalId ="";


Office.onReady((info) => {


  if (info.host === Office.HostType.Outlook) {

    Entorno = new EntornoClase();

  };

});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Mi Dominio",
    icon: "Icon.32x32",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
};

async function Opcion1(event) {
  Entorno.OutlookCorreoActual = Office.context.mailbox.item;
  await Entorno.OpcionA1();
  event.completed();
};


async function prependHeaderOnSend(event) {

  Entorno.OutlookCorreoActual = Office.context.mailbox.item;
  await Entorno.Correo.SetDatosGenerarTexto();
  if (TokenCliente) {
    Office.context.mailbox.item.body.getTypeAsync(
      {
        asyncContext: event
      },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          if (asyncResult.error) {
            Entorno.MostrarMensaje("No se pudo pegar el texto", "generartexto");
          } else {
            Entorno.MostrarMensaje("No se pudo pegar el texto", "generartexto");
          }
          
          // Si la operación falla, establecer un intervalo para volver a intentarlo cada minuto
          intervalId = setInterval(() => {
            prependHeaderOnSend(event);
          }, 60000);
  
          return;
        }
  
        // Si la operación tiene éxito, pegar el texto y detener el intervalo
        clearInterval(intervalId);
        Entorno.PegarTexto(asyncResult); 
  
    });

  }else{

    //Identificarse
  };
}




Office.actions.associate("prependHeaderOnSend", prependHeaderOnSend);





class MiDominioClase {

   async GetTextoExtension() {

    Entorno.OutlookCorreoActual = Office.context.mailbox.item;

    await Entorno.Correo.SetDatosGenerarTexto();

      let Resultado = Entorno.Correo.Asunto;
      return Resultado;
  }

};
class CorreoClase {

  constructor() {
      this.Asunto = "";
      this.CuerpoHTML = "";
      this.Texto = "";
  };

  
  async SetDatosGenerarTexto(){

    try {

      Entorno.Correo.Asunto = "";
      Entorno.Correo.CuerpoHTML = "";
                             
      Entorno.Correo.Asunto = await new Promise((resolve, reject) => {
                                                  Entorno.OutlookCorreoActual.subject.getAsync((result) => {
                                                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                                                      resolve(result.value);
                                                    } else {
                                                      reject('');
                                                    }
                                                  });
                                                });
  
    } catch (error) {
      Entorno.MostrarMensaje("no se pudo pegar el texto", "generartexto");
    }

  };



};
class EntornoClase {

   
  constructor() {

      this.OutlookCorreoActual = null;
      this.Correo = new CorreoClase();
      this.MiDominio = new MiDominioClase();

  };


  async MostrarMensaje(mensajeMostrar, accionid) {
    const message = {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: mensajeMostrar + "  ",
      icon: "Icon.32x32",
      persistent: true
    };
    Office.context.mailbox.item.notificationMessages.replaceAsync(accionid, message);  
   
  };
  async MostrarMensajeIcono(mensajeMostrar, accionid, iconoMostrar) {
    const message = {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: mensajeMostrar + "  ",
      icon: iconoMostrar,
      persistent: true
    };
    Office.context.mailbox.item.notificationMessages.replaceAsync(accionid, message);  
   
  };
  async CerrarMensaje(accionid) {
    Office.context.mailbox.item.notificationMessages.removeAsync(accionid); 
  }
  



  async CorregirTexto(Texto) {

    var textArea = document.createElement('textarea');
    textArea.innerHTML = Texto;
    return textArea.value;

  };






    

  async OpcionA1(){

      Entorno.MostrarMensajeIcono("Opcion A1" , "opciona1", "icon.naranja");      

  };
  
  async PegarTexto(Datos){

    Entorno.MostrarMensaje("Generando texto", "generandotexto");



    const bodyFormat = Datos.value;
    let texto = await Entorno.MiDominio.GetTextoExtension();
    let textoFinal = await Entorno.CorregirTexto(texto);
    Entorno.CerrarMensaje("generandotexto");

    if (textoFinal ==''){
      Entorno.MostrarMensaje("No se pudp pegar el texto", "generartextorespuesta");
      return;
    }else{
      console.log("Office.context.mailbox.item.body",Office.context.mailbox.item.body)
      Office.context.mailbox.item.body.prependOnSendAsync(
        textoFinal,
        {
          asyncContext: Datos.asyncContext,
          coercionType: bodyFormat
        },
        (Datos) => {
          console.log("Datos",Datos)
          if (Datos.status === Office.AsyncResultStatus.Failed) {
            Entorno.MostrarMensaje("No se pudp pegar el texto", "generartextorespuesta");
            return;
          }
           Datos.asyncContext.completed();          
        }
      );
      Entorno.MostrarMensaje("Texto generado", "generartextorespuesta");

    };


  };








};

var Entorno = null;




 function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}
const g = getGlobal();

g.action = action;
g.Opcion1 = Opcion1;
g.Entorno = Entorno;












