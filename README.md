Readme
=======
**Proyecto de Mejoramiento cuenta de Cobro**
_______________________________

Las cuentas de cobro son un documento legal requerido para poder hacer la solicitud 
de un pago a una entidad, empresa o persona. 
Esta debe cumplir algunos requisitos mínimos para que sea aceptada como un soporte 
contable.

1. Antecedentes

Inicialmente las cuentas de cobro se realizaban a puño y letra de cada persona, 
luego pasaron a realizarse en máquinas de escribir y después se pudo automatizar 
un poco teniendo en una hoja electrónica los datos bien estructurados de la información
 que se necesitaba. Si se necesita generar muchas cuentas de cobro de diferentes 
 personas se va a necesitar hacer uso de un formato especifico para mostrar todos 
 los datos de una manera ordenada, clara y exacta. 
Cuando se hacían a mano no se entendía la letra o los números presentaban confusión
 por la caligrafía de cada persona, por lo cual con la llegada del computador y la
 estandarización de procesos se pueden manejar tipografías de letra entendibles. 

2. Inconvenientes con una solución en Excel

Inicialmente, el proceso fue realizado en combinación de una hoja de Excel que 
contenía todos los datos con un Word en una combinación de correspondencia, 
sin embargo esto no cumplía un requisito importante y es que debe aparecer el valor
 no solo en números sino también en letras, y escribir cada valor en letras a veces
 tiende a generar confusiones o inexactitudes y es dispendioso. 
 
## Se procede a realizar una primera solución con dos macros en visual Basic:

Una que permita pasar el valor de números a letras y la otra que genere cada cuenta
 de cobro en pdf en forma individual.

Sabemos que el Excel posee una capacidad limitada de registros y se vuelve lento la
 ejecución de una macro si esta cantidad aumenta.
Se utilizan funciones de búsqueda para ubicar la información y requiere tener siempre
 el mismo orden de las columnas, si se adiciona una nueva el sistema falla.

** Solución propuesta en Python
Se propone realizar el mismo proceso en Python teniendo la información estructurada
 en Excel y arrojando un pdf por cada registro.
Se deben tener dos funciones:

Una que convierta los números a letras y otra que genere el pdf.

## Datos de la hoja electrónica
Se consideran los siguientes datos en la hoja electrónica:
•	#Cuenta_Cobro
•	Nombre_del_artista
•	Fecha
•	Nombre Empresa y Nit 
•	Valor(números)
•	Concepto
•	Identificación del artista
•	Forma de pago
Estos son los datos mínimos que se requerirían en la cuenta de cobro puede haber más
 según los requisitos de la entidad solicitante.
Por cada registro se debe generar un pdf

## Formato de la cuenta de cobro básico:

Fecha
Numero cuenta de cobro# 

Entidad que pagara 
Nit identificación

DEBE A:

Nombre del artista
Identificación del artista


Valor de (números) {valor en letras}

Por el concepto de: {concepto} & {nombre del evento}


Cordialmente,

Nombre del artista
Identificación del artista



## Ejecución del script en python
Se requiere tener instalados las siguientes librerías:
import pandas as pd
import pdfkit
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import tkinter as tk
from tkinter import simpledialog
import pythoncom
import winshell
import time

# Consideraciones a tener en cuenta:
•	Corre en Windows
•	Se requiere un Excel con los datos mínimos de la cuenta de cobro
•	Genera un shorcut en el escritorio
•	El formato de la cuenta esta diseñado en tamaño carta y posee en el script 
las márgenes de 4 cms a la derecha e izquierda (113.4 p), margen superior 
de 3 cm (85.04 p) e inferior de 2 cm ( 56.69 p)
•	Creara una carpeta llamada cuentas para ubicar los pdfs
•	Preguntara por un numero de registro inicial y uno final

##Modificaciones de mejora

Dentro del script se solicita crear una carpeta para ubicar los pdf.
Se crea un input que pregunte al usuario cual cuenta de cobro desea generar, 
ya sea una o un rango desde un numero inicial a uno final. 
Esta interacción con el usuario final se dará con shortcut en el escritorio para 
facilidad de correr el sistema.
