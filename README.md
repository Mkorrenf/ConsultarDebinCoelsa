# ConsultarDebinCoelsa
Aplicación para consultar el detalle de Debin de forma individual o masiva a la cámara compensadora Coelsa

•	Para el correcto funcionamiento de la aplicación, se requiere tener el correspondiente certificado generado por la cámara, ya que, si no va a ser rechazada la conexión.

•	Es necesario modificar el archivo App.config, los parámetros de "baseAddress" y "urlConsultaDebin" con los valores correspondientes al endpoint a utilizar.

•	Los parámetros "rutaDefectoGuardarArchivo" y "rutaDefectoCardarArchivo", permiten configurar valores por defecto al momento de cargar la aplicación, para simplificar el uso.

Las consultas, tanto individuales, como masivas, se van a ir almacenando y acumulando en un datatable y a su vez se van a ir mostrando en el datagridview. Una vez finalizada todas las consultas requeridas, se pueden exportar a un archivo Excel. En el caso de ser necesario se pueden limpiar las consultas para volver a empezar.

Cualquier comentario o aporte serán bienvenidos.
