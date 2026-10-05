Este repositorio contiene el archivo "Option Explicit.bas" para reajustar valores a negativos en 
estados de cuentas con el software de facturacion MemoryG2000 en version Figaro (Desktop/Escritorio) para Windows.
Extensiones ODS, XLS.

Instalacion de la Macro:
1) Ingresamos a OpenOffice Calc en Windows.
2) Vamos a Herramientas > Macros > Organizar Macros > OpenOffice Basic...
3) En el cuadro "Macros Basic de OpenOffice" veremos que estamos en Standard > Module1
4) La parte derecha tenemos varias opciones, elegimos unicamente: "Editar".
5) Copias todo el contenido del archivo descargado "Option Explicit.bas" y lo pegamos
en el nuevo recuadro en su campo de texto, sin antes haber vaciado/eliminado todo texto/codigo ya existente.
6) Le damos a Ctrl+G para guardar cambios.
7) Primero comprobamos que funciona: Si estamos justo con el archivo del estado de cuenta, antes de cerrar la ventana presionamos en el teclado "F5"
o al botón Ejecutar Programa BASIC ubicado al lado del botón Compilar; Pueden verificar el nombre del boton colocando el mouse encima y saber cual es.
8) Una vez guardado y confirmado el funcionamiento (de no confirmar continuamos si actúa luego):
  Vamos a Herramientas > Macros > Organizar Macros > OpenOffice Basic...
  Seleccionamos "Module1" (en la parte izquierda) y luego damos click al botón "Asignar..." (parte derecha).
  Vamos a la pestaña "Eventos" (parte superior de la ventana) y en la lista elejimos:
  "La carga del documento ha finalizado." y hacemos click en el botón "Macro", ahi elegimos Module1 y guardamos.
  Reiteramos la accion pero seleccionamos en la lista "Abrir Documento", esto para asegurar la carga del script si hay errores.

NOTA: Requiera Java 8 (JRE) o superior indicado por OpenOffice para proceder con Macros.

Con esto hecho, al abrir un EC (Estado de cuenta) como archivo ODS/XLS impreso por Memory G2000 Figaro,
veremos que se recalcula los saldos donde Devoluciones y Notas de Credito pasan como valor negativo.
Esto para agilizar el trabajo y no depender manualmente de cambiar valores enteros positivos a negativos en facturas que sean Devoluciones/Notas de Crédito.
