# ConsultarDebinCoelsa
Aplicación para consultar el detalle de Debin de forma individual o masiva a la cámara compensadora Coelsa.

Para el correcto funcionamiento de la aplicación:
•	Se requiere tener el correspondiente certificado generado por la cámara, ya que, si no, va a ser rechazada la conexión.

•	Es necesario modificar el archivo App.config, los parámetros de "baseAddress" y "urlConsultaDebin" con los valores correspondientes al endpoint a utilizar.

•	Los parámetros "rutaDefectoGuardarArchivo" y "rutaDefectoCardarArchivo", permiten configurar valores por defecto al momento de cargar la aplicación, para simplificar el uso.

Composición y Funcionamiento:
La aplicación es en formato escritorio, con Winforms en lenguaje C#. Se consume mediante HttpClient el servicio de consulta de debin de Coelsa. Con la respuesta obtenida, se deserializa del formato json mediante Newtonsoft, a clases creadas para poder ir incorporando los registros en el datatable. También, esto permite, manipular los datos en caso de ser necesario. Como, por ejemplo, es posible persistir en una BD propia, actualizar el estado de operaciones, o lo que sea necesario.

Las consultas, tanto individuales, como masivas, se van a ir almacenando y acumulando en un datatable y a su vez se van a ir mostrando en el datagridview. Una vez finalizada todas las consultas requeridas, se pueden exportar a un archivo Excel. En el caso de ser necesario se pueden limpiar las consultas para volver a empezar.

Cualquier comentario o aporte serán bienvenidos.
