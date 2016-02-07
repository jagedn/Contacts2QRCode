# Contacts2QRCode
AddOn for Google Sheets


Contacts2QRCode es un AddOn para Google Sheet que permite generar códigos QR de contacto utilizando la información de tu libreta de correo.
Para usarlo, simplemente instalalo en una hoja de google y utiliza el diálogo que proporciona y que te irá giando en el proceso:
- seleccionar un grupo de contactos de tu agenda
- extraer y volcar la información de los miembros del grupo
- ayudarte en la edición de los campos comunes (usar una misma dirección para todos, url de contacto, etc)
- generar en tu Drive las imágnes correspondientes para que puedas imprimirlas, usarlas con el móvil, etc


# Codigo
Code.gsp contiene la parte de código que corre en el servidor. Es la encargada de acceder a tu cuenta, leer los contactos, etc.

El resto corresponden a la parte de código que corre en el navegador del usuario y que llaman a las funciones del servidor.
