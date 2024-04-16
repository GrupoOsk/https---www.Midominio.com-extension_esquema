/* global document, Office, g */


Office.onReady((info) => {

  if (info.host === Office.HostType.Outlook) {



      PrimeraVez = false;
      // document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "flex";
      
      $(function() { 

        $(document).on('click', '.opcionmenu_contenedor', function (event) {
            event.preventDefault();
            let accionSeleccionada = $(this).data("accion");
            opcionMenuAccion(accionSeleccionada);
        });    
        $(document).on('click', '.panel_body_2_volver', function (event) {
            event.preventDefault();
            $(".panel_body_inicial").show();
            $(".panel_body_2").hide();
        });  
        




      });



  };

});


async function opcionMenuAccion(accionSeleccionada){



  switch (accionSeleccionada) {

   
    case "accion_2":
        $(".panel_body_inicial").hide();
        $(".panel_body_2").show();
        break;

    default:
        await g.Entorno.MostrarMensaje("Opción seleccionada incorrecta.");
        break;
  };


};



