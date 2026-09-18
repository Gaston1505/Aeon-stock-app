# Extrae los precios de la competencia de "Precios al por menor.xlsx" (5 hojas: AA, Anafe,
# Campana, Termocalefones, Horno) y arma un batch de Firestore REST (writes[]) listo para
# hacer commit contra preciosMercado. No requiere Node/Python — usa Excel COM (ya instalado)
# para leer el archivo, y curl (en Bash) hace el POST real.

$ErrorActionPreference = "Stop"
$srcPath = "C:\Users\Quantum-I\Desktop\AEON\Precios\Competencia al por Menor\Precios al por menor.xlsx"
$outPath = "C:\Users\Quantum-I\Desktop\AEON\APP\scripts\precios_mercado_batch.json"

# specCols = cantidad de columnas de "ficha técnica" antes de que empiecen los grupos de
# fuente (cada grupo son 3 columnas: Fecha, Precio Gs, Precio U$S, identificadas por la celda
# combinada de la fila 1). marcaCol = columna (1-based) con el fabricante, si la hoja la trae
# separada; si no, se deja vacía.
$hojas = @(
  @{ nombre = "AA";             categoriaPrincipal = "Aire Acondicionado"; subcategoria = "";        specCols = 3; marcaCol = 0 },
  @{ nombre = "Anafe";          categoriaPrincipal = "Cocina";             subcategoria = "Anafe";    specCols = 5; marcaCol = 5 },
  @{ nombre = "Campana";        categoriaPrincipal = "Cocina";             subcategoria = "Campana";  specCols = 8; marcaCol = 7 },
  @{ nombre = "Termocalefones"; categoriaPrincipal = "Termocalefones";     subcategoria = "";         specCols = 6; marcaCol = 0 },
  @{ nombre = "Horno";          categoriaPrincipal = "Cocina";             subcategoria = "Horno";    specCols = 6; marcaCol = 6 }
)

function Parse-FechaISO($cell) {
  $v2 = $cell.Value2
  if ($v2 -is [double]) {
    $dt = [DateTime]::FromOADate($v2)
    return $dt.ToString("yyyy-MM-dd")
  }
  $t = "$v2".Trim()
  if ($t -eq "") { return $null }
  # Typo conocido: "13/082026" -> "13/08/2026"
  if ($t -match '^(\d{1,2})\/(\d{2})(\d{4})$') {
    $t = "$($Matches[1])/$($Matches[2])/$($Matches[3])"
  }
  try {
    $dt = [DateTime]::ParseExact($t, "dd/MM/yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
    return $dt.ToString("yyyy-MM-dd")
  } catch {
    return $null
  }
}

function Num($cell) {
  $v2 = $cell.Value2
  if ($v2 -is [double] -or $v2 -is [int]) { return [double]$v2 }
  return 0
}

function Texto($cell) {
  $t = $cell.Text
  if ($null -eq $t) { return "" }
  return "$t".Trim()
}

# Normaliza mayusculas/minusculas de nombres de tienda/marca (la planilla mezcla "Bristol",
# "bristol", "OLIER" segun la hoja): solo pone mayuscula la primera letra de cada palabra,
# sin tocar el resto -- asi una sigla que ya vino en mayuscula (MIDEA, RHEEM) queda intacta.
function TituloSuave($s) {
  if ($s -eq "") { return $s }
  $palabras = $s -split '(\s|/)'
  $resultado = @()
  foreach ($p in $palabras) {
    if ($p.Length -eq 0 -or $p -eq " " -or $p -eq "/") { $resultado += $p; continue }
    $resultado += $p.Substring(0,1).ToUpperInvariant() + $p.Substring(1)
  }
  return ($resultado -join "")
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($srcPath)

$registros = @()

foreach ($h in $hojas) {
  $ws = $wb.Sheets.Item($h.nombre)
  $ur = $ws.UsedRange
  $rows = $ur.Rows.Count
  $cols = $ur.Columns.Count

  # Cabecera fila 2 (nombres de columnas de ficha técnica, para armar "especificacion")
  $specHeaders = @()
  for ($c = 1; $c -le $h.specCols; $c++) {
    $specHeaders += (Texto $ws.Cells.Item(2, $c))
  }

  # Grupos de fuente: cada 3 columnas desde specCols+1, nombre = celda combinada fila 1
  $grupos = @()
  $c = $h.specCols + 1
  while ($c -le $cols) {
    $nombreFuente = Texto $ws.Cells.Item(1, $c)
    if ($nombreFuente -ne "") { $grupos += @{ col = $c; fuente = $nombreFuente } }
    $c += 3
  }

  for ($r = 3; $r -le $rows; $r++) {
    # especificacion: junta "header: valor" de las columnas de ficha tecnica (menos marca, que va aparte)
    # Separador ASCII a proposito: Windows PowerShell 5.1 lee un .ps1 sin BOM con el codepage
    # del sistema, asi que un caracter no-ASCII literal en el propio script (como "*") se
    # corrompe al leerlo -- mejor evitarlo aca y dejar los acentos solo en los VALORES que
    # vienen del xlsx (esos se leen bien via COM, no via el codigo fuente del script).
    $partes = @()
    $marca = ""
    for ($sc = 1; $sc -le $h.specCols; $sc++) {
      $val = Texto $ws.Cells.Item($r, $sc)
      if ($val -eq "") { continue }
      if ($h.marcaCol -gt 0 -and $sc -eq $h.marcaCol) { $marca = $val; continue }
      $label = $specHeaders[$sc - 1]
      if ($label -ne "") { $partes += "$label $val" } else { $partes += $val }
    }
    if ($partes.Count -eq 0) { continue }
    $especificacion = ($partes -join " - ")

    foreach ($g in $grupos) {
      $precioGs = Num $ws.Cells.Item($r, $g.col + 1)
      $precioUsd = Num $ws.Cells.Item($r, $g.col + 2)
      if ($precioGs -le 0 -and $precioUsd -le 0) { continue }
      $fecha = Parse-FechaISO $ws.Cells.Item($r, $g.col)
      if ($null -eq $fecha) { $fecha = "" }
      $registros += [PSCustomObject]@{
        categoriaPrincipal = $h.categoriaPrincipal
        subcategoria       = $h.subcategoria
        especificacion     = $especificacion
        marca              = TituloSuave $marca
        empresa            = TituloSuave $g.fuente
        fecha              = $fecha
        precioGs           = [math]::Round($precioGs)
        precioUsd          = [math]::Round($precioUsd)
        fuenteTipo          = "Planilla propia"
        notas              = ""
        productoEquivalenteId = ""
      }
    }
  }
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Output "Registros extraidos: $($registros.Count)"
$registros | ConvertTo-Json -Depth 5 | Out-File -FilePath $outPath -Encoding utf8
Write-Output "Guardado en $outPath"

# Arma el payload para POST .../documents:commit (batchWrite de hasta 500 escrituras en un
# solo request) -- mucho mas rapido y menos propenso a mojibake que 49 curls sueltos.
$proyecto = "aeon-stock-app"
$baseName = "projects/$proyecto/databases/(default)/documents/preciosMercado"
$createdAt = [int64]([DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds())
$rand = New-Object System.Random

function NuevoId() {
  $chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
  -join (1..20 | ForEach-Object { $chars[$rand.Next(0, $chars.Length)] })
}

function StrVal($s) { return @{ stringValue = "$s" } }
function IntVal($n) { return @{ integerValue = "$([int64]$n)" } }

$writes = @()
foreach ($reg in $registros) {
  $fields = @{
    categoriaPrincipal    = StrVal $reg.categoriaPrincipal
    subcategoria          = StrVal $reg.subcategoria
    especificacion        = StrVal $reg.especificacion
    marca                 = StrVal $reg.marca
    empresa               = StrVal $reg.empresa
    fecha                 = StrVal $reg.fecha
    precioGs              = IntVal $reg.precioGs
    precioUsd             = IntVal $reg.precioUsd
    fuenteTipo            = StrVal $reg.fuenteTipo
    notas                 = StrVal $reg.notas
    productoEquivalenteId = StrVal $reg.productoEquivalenteId
    createdAt             = IntVal $createdAt
  }
  $writes += @{ update = @{ name = "$baseName/$(NuevoId)"; fields = $fields } }
}

$commitPath = "C:\Users\Quantum-I\Desktop\AEON\APP\scripts\precios_mercado_commit.json"
(@{ writes = $writes } | ConvertTo-Json -Depth 8) | Out-File -FilePath $commitPath -Encoding utf8
Write-Output "Payload de commit guardado en $commitPath ($($writes.Count) writes)"
