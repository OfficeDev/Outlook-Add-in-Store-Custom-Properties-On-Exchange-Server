---
page_type: sample
products:
- office-outlook
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 8/19/2015 11:13:31 AM
---
# Outlook-Add-in-Store-Custom-Properties-On-Exchange-Server
En este ejemplo, se muestra cómo establecer una propiedad en un mensaje de correo electrónico y después almacenar esa propiedad en su servidor Exchange para recuperarla la próxima vez que se devuelva el elemento. Por ejemplo, si el complemento de correo para Outlook agrega contactos a una base de datos de contactos externa, puede establecer una propiedad en un elemento para indicar que se ha agregado un contacto, por lo que no se le pide que agregue el mismo contacto por segunda vez.

El método [loadCustomPropertiesAsync](http://msdn.microsoft.com/library/dfbec151-8ea7-4915-b723-09ea1396a261) del objeto del elemento devuelve un objeto [CustomProperties](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3) que contiene y administra las propiedades personalizadas que ha guardado para un elemento. Después de cargar las propiedades personalizadas, puede hacer lo siguiente:

* Usar los métodos [get](http://msdn.microsoft.com/library/3ab90551-138a-482d-9d93-4cdb20db193b) y [set](http://msdn.microsoft.com/library/03a8b253-b681-4a09-b828-80d9cf46ca9d) para leer y escribir propiedades personalizadas. 
* Usar el método [remove](http://msdn.microsoft.com/library/01983beb-766f-4308-9e23-e840e950f7e3) para eliminar las propiedades personalizadas que creó. 
* Usar el método [saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56) para guardar los cambios que haya realizado en el servidor de Exchange. 

Debe llamar al método [saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56) para almacenar las propiedades en el servidor de Exchange; en caso contrario, todos los cambios realizados se descartan cuando se cambia el elemento actual.

La interfaz de usuario de ejemplo tiene tres páginas: una para establecer la clave y el valor de una propiedad personalizada, otra para recuperar el valor de una propiedad personalizada y otra para quitar las propiedades personalizadas o para guardar los cambios que realice en el servidor de Exchange.

El archivo JavaScript contiene controladores de clic para botones en la interfaz de usuario para obtener, establecer, quitar y guardar propiedades personalizadas mediante los métodos correspondientes del objeto [CustomProperties](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3). Se establece una variable booleana local, customPropertiesAreLoaded, en la función callback para que el método loadCustomPropertiesAsync muestre que el objeto Custom Properties está cargado. El controlador comprueba este valor para asegurarse de que el objeto [CustomProperties](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3) está disponible antes de llamar a las funciones en el objeto. 

*Requisitos previos*

Para este ejemplo se necesita lo siguiente:

* Visual Studio 2012, con las aplicaciones para plantillas de proyecto Office. 
* Un equipo que ejecute Exchange 2013 y, como mínimo, una cuenta de correo electrónico o una cuenta de desarrollador de Office 365. Puede [participar en el programa para desarrolladores Office 365 y obtener una suscripción gratuita durante 1 año a Office 365](https://aka.ms/devprogramsignup).
* Familiaridad con los servicios web y la programación de JavaScript. 
* Internet Explorer 9 o la Versión preliminar de Internet Explorer 10. 

*Componentes clave del ejemplo*

La solución de ejemplo contiene los archivos siguientes:

* Proyecto CustomProperties 
  * CustomProperties.xml: el archivo de manifiesto para el complemento de correo de Outlook. 
* Proyecto CustomPropertiesWeb
  * Home.html: la interfaz de usuario HTML para el complemento de correo electrónico de Outlook. 
  * Home.js: es el archivo JavaScript que controla la solicitud y el uso de la solicitud de servicios Web Exchange (EWS). 
  * Scripts\Lib: el complemento de correo para Outlook y la API de Outlook Web App. 


*Configuración de la muestra*

El complemento de correo se activará en cualquier mensaje de correo electrónico de la bandeja de entrada del usuario. Puede hacer que sea más fácil probar el complemento enviando uno o más mensajes de correo electrónico a la cuenta de prueba antes de ejecutar el ejemplo.

*Compile el ejemplo*

Pulse F5 para compilar e implementar la aplicación de ejemplo. Para implementar la aplicación complete las siguientes tareas:

1. Conecte a una cuenta de Exchange proporcionando la dirección de correo electrónico y la contraseña de un servidor de Exchange 2013. 
2. Permita que el servidor configure la cuenta de correo electrónico. 

*Inicie y pruebe la aplicación*

Ejecute y pruebe el ejemplo en el explorador web que inicia Visual Studio cuando compile e implemente el ejemplo.

Si está ejecutando el ejemplo en un servidor de Exchange que usa el certificado autofirmado predeterminado, recibirá un error de certificado cuando se inicie el explorador web. Después de que haya comprobado que el explorador web está abriendo la dirección URL correcta en la dirección web, puede hacer clic en Continuar a este sitio web para iniciar Outlook Web App.

Siga estos pasos para ejecutar la muestra:

1. Inicie sesión en la cuenta de correo electrónico escribiendo el nombre de la cuenta y la contraseña. 
2. Seleccione un mensaje en la Bandeja de entrada 
3. Espere a que aparezca la barra de la aplicación sobre el mensaje. 
4. En la barra de la aplicación, haga clic en Propiedades personalizadas. 
5. Cuando aparezca el complemento de correo de Propiedades personalizadas, escriba el nombre y el valor de la propiedad en los cuadros de texto y después haga clic en el botón Guardar para guardar el valor de la propiedad. 
6. Haga clic en Obtener, escriba el nombre de una propiedad y después haga clic en el botón Obtener para recuperar una propiedad. 
7. Haga clic en Administrar y después en Conservar para guardar las propiedades almacenadas en el servidor de Exchange, o bien escriba un nombre de propiedad y haga clic en Quitar para eliminar la propiedad del almacenamiento. 

*Solución de problemas*

A continuación se enumeran los errores comunes que pueden ocurrir al usar Outlook Web App para probar un complemento de correo para Outlook:

* La barra de la aplicación no aparece cuando se selecciona un mensaje. Si esto ocurre, vuelva a iniciar la aplicación seleccionando Depuración: detener depuración en la ventana de Visual Studio y presione F5 para recompilar e implementar la aplicación. 
* Es posible que los cambios en el código de JavaScript no se hayan recogido al implementar y ejecutar la aplicación. Si no se han añadido los cambios, borre la memoria caché en el explorador web. Para ello, seleccione Herramientas: opciones de Internet y haga clic en el botón Eliminar.... Elimine los archivos temporales de Internet y reinicie la aplicación. 

*Recursos adicionales*

* [Más complementos de ejemplo](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [CustomProperties (objeto)](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3)
* [loadCustomPropertiesAsync method](http://msdn.microsoft.com/library/dfbec151-8ea7-4915-b723-09ea1396a261)
* [Método get](http://msdn.microsoft.com/library/3ab90551-138a-482d-9d93-4cdb20db193b)
* [Método set](http://msdn.microsoft.com/library/03a8b253-b681-4a09-b828-80d9cf46ca9d)
* [Método remove](http://msdn.microsoft.com/library/01983beb-766f-4308-9e23-e840e950f7e3)
* [Método saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56)



Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, vea [Preguntas frecuentes sobre el código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.
