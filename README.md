# QR Label Maker (VB WinForms)

Este repositorio contiene un proyecto de ejemplo (Windows Forms en Visual Basic) que lee un archivo Excel (.xlsx), genera códigos QR por fila, compone etiquetas (QR + texto) y exporta las etiquetas como imágenes PNG individuales y hojas con varias etiquetas por imagen.

Contenido agregado:
- Program.vb - punto de entrada
- Form1.vb - formulario WinForms con UI y lógica para generar etiquetas
- setup_project_final.ps1 - script PowerShell que crea/ensambla el proyecto localmente y empaqueta un ZIP

Requisitos
- .NET SDK (dotnet)
- Windows (WinForms)
- Visual Studio 2022/2019 o dotnet CLI

Instrucciones rápidas
1. Clona este repositorio o descarga los archivos.
2. Ejecuta setup_project_final.ps1 en PowerShell (opcionalmente pasando una carpeta destino):
   powershell -ExecutionPolicy Bypass -File .\setup_project_final.ps1
   Esto creará un proyecto WinForms VB en la carpeta destino y generará QrLabelMaker.zip.
3. Abre el proyecto en Visual Studio y ejecuta.

Uso de la aplicación
- Selecciona el archivo Excel (.xlsx) con la primera fila como encabezados.
- Indica las columnas que usarás para el contenido del QR (coma-separadas) y las columnas a mostrar en la etiqueta.
- Haz "Vista previa" para ver una etiqueta de ejemplo.
- Haz "Generar etiquetas" para exportar PNGs a la carpeta seleccionada.

Licencias y notas
- EPPlus se usa en modo LicenseContext.NonCommercial en el código. Revisa la licencia si vas a usarlo comercialmente.
- QRCoder es una librería de generación de QR en .NET.
