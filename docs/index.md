## Repositorio de macros para LibreOffice

En este repositorio se irán subiendo macros para LibreOffice.

Las macros disponibles en este momento son:
- InitialEvaluation2

### InitialEvaluation2
Con esta macro se pretende generar el boletín de notas de cada uno de los alumnos de un grupo a partir de contenido de una hoja de cálculo.

Para utilizar esta macro, se ha de insertar el contenido del fichero InitalEvaluation2.bas como una macro dentro de la hoja de cálculo InitialEvaluation2.ods. En la parte superior de la hoja de cálculo hay que poner el **Grupo**, la **Fecha firma** y el **Tutor** correspondientes. En la columna **Alumno** hay que poner el nombre y apellidos de cada uno de los alumnos.

La función principal de esta macro (la que hay que ejecutar) es *makeEvalDocument2*

Para el correcto funcionamiento de la macro es necesario que el archivo *template2IE.odt* esté en la misma carpeta que el archivo de la hoja de cálculo. Este archivo contiene la plantilla del documento. La plantilla se puede personalizar cambiando los logos de la cabecera, así como el pie de página.

El documento creado por la macro se guardará también en la misma carpeta.
