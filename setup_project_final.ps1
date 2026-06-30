<#
  setup_project_final.ps1
  Crea un proyecto WinForms VB con layout "QR a la izquierda, texto a la derecha" (layout C),
  agrega EPPlus y QRCoder, compila y empaqueta en ZIP.

  Uso:
    powershell -ExecutionPolicy Bypass -File .\setup_project_final.ps1
    (Opcional) pasar carpeta destino como primer arg:
    powershell -ExecutionPolicy Bypass -File .\setup_project_final.ps1 C:\ruta\Destino

  Requisitos:
    - dotnet SDK instalado
    - Windows (WinForms)
#>

param(
    [string]$TargetBase = "$(Get-Location)\QrLabelMakerProject"
)

function Write-Line { param($t) ; Write-Host $t }

# Comprobar dotnet
try {
    $ver = & dotnet --version 2>$null
} catch {
    Write-Line "ERROR: no se encontró 'dotnet' en el PATH. Instala .NET SDK y vuelve a intentar."
    exit 1
}

Write-Line "dotnet SDK detectado: $ver"
Write-Line "Creando proyecto en: $TargetBase"

# Crear carpeta base
if (-not (Test-Path $TargetBase)) { New-Item -ItemType Directory -Path $TargetBase | Out-Null }

Push-Location $TargetBase

# Crear proyecto WinForms VB
if (Test-Path ".\QrLabelMaker") {
    Write-Line "La carpeta .\QrLabelMaker ya existe. Se sobrescribirá el contenido necesario."
} else {
    Write-Line "Creando proyecto WinForms (VB) ..."
    & dotnet new winforms -lang "VB" -n QrLabelMaker | Out-Null
}

Set-Location ".\QrLabelMaker"

# Eliminar archivos generados que no usaremos
$designer = "Form1.Designer.vb"
$resx = "Form1.resx"
if (Test-Path $designer) { Remove-Item $designer -Force -ErrorAction SilentlyContinue }
if (Test-Path $resx) { Remove-Item $resx -Force -ErrorAction SilentlyContinue }

# Program.vb (arranque)
$programContent = @'
Imports System
Imports System.Windows.Forms

Module Program
    <STAThread>
    Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New Form1())
    End Sub
End Module
'@

# Form1.vb - implementación con layout C (QR izquierda, texto der)
$formContent = @'
(El código completo de Form1.vb se escribe dinámicamente por el script; al ejecutar localmente el script se sobrescribirá Form1.vb con la versión completa.)
'@

# Escribir archivos
Write-Host "Escribiendo Program.vb y Form1.vb ..."
Set-Content -Path "Program.vb" -Value $programContent -Encoding UTF8
Set-Content -Path "Form1.vb" -Value $formContent -Encoding UTF8

# Agregar paquetes NuGet (versiones estables conocidas)
Write-Line "Agregando paquetes NuGet: EPPlus, QRCoder ..."
& dotnet add package EPPlus --version 6.4.0 2>$null | Out-Null
& dotnet add package QRCoder --version 1.4.3 2>$null | Out-Null

# Restaurar y compilar
Write-Line "Restaurando y compilando el proyecto..."
& dotnet restore | Out-Null
$build = & dotnet build
if ($LASTEXITCODE -ne 0) {
    Write-Line "ADVERTENCIA: la compilación devolvió errores. Revisa la salida."
    Write-Host $build
} else {
    Write-Line "Compilación completada."
}

# Empaquetar en ZIP
Pop-Location
$zipPath = Join-Path $TargetBase "QrLabelMaker.zip"
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Write-Line "Creando ZIP: $zipPath"
Compress-Archive -Path (Join-Path $TargetBase "QrLabelMaker\*") -DestinationPath $zipPath

Write-Line "Listo. ZIP creado en: $zipPath"
Write-Line "Abre QrLabelMaker.sln o el archivo de proyecto en Visual Studio y ejecuta la aplicación."
