<#
Perspective Editor local proxy v3.
Uses TOM (Microsoft.AnalysisServices.Tabular) for writes; DMVs for reads.

Usage:
  pwsh -File perspective-server-v3.ps1
  Listens on http://localhost:8768
#>

$ErrorActionPreference = 'Stop'

$pbiBin = "C:\Program Files\Microsoft Power BI Desktop\bin"
$adomdPath  = Join-Path $pbiBin "Microsoft.PowerBI.AdomdClient.dll"
$tomCore    = Join-Path $pbiBin "Microsoft.AnalysisServices.Server.Core.dll"
$tomTabular = Join-Path $pbiBin "Microsoft.AnalysisServices.Server.Tabular.dll"

foreach ($p in @($adomdPath, $tomCore, $tomTabular)) {
  if (-not (Test-Path $p)) { Write-Error "Missing $p"; exit 1 }
}

[void][Reflection.Assembly]::LoadFrom($adomdPath)
[void][Reflection.Assembly]::LoadFrom($tomCore)
[void][Reflection.Assembly]::LoadFrom($tomTabular)

function New-AsConnection([int]$port, [string]$catalog) {
  $cs = "Data Source=localhost:$port"
  if ($catalog) { $cs += ";Catalog=$catalog" }
  $c = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdConnection $cs
  $c.Open()
  return $c
}

function Invoke-Query([object]$conn, [string]$sql) {
  $cmd = $conn.CreateCommand()
  $cmd.CommandText = $sql
  $rdr = $cmd.ExecuteReader()
  $cols = @(); for ($i=0; $i -lt $rdr.FieldCount; $i++) { $cols += $rdr.GetName($i) }
  $rows = @()
  while ($rdr.Read()) {
    $o = [ordered]@{}
    for ($i=0; $i -lt $rdr.FieldCount; $i++) { $o[$cols[$i]] = $rdr[$i] }
    $rows += [PSCustomObject]$o
  }
  $rdr.Close()
  return $rows
}

function Get-Catalog([int]$port) {
  $conn = New-AsConnection $port $null
  try {
    $cat = Invoke-Query $conn "SELECT [CATALOG_NAME] FROM `$SYSTEM.DBSCHEMA_CATALOGS"
    if (-not $cat -or $cat.Count -eq 0) { throw "No catalog on port $port" }
    return [string]$cat[0].CATALOG_NAME
  } finally { $conn.Close() }
}

function Get-Model([int]$port) {
  $catalog = Get-Catalog $port
  $conn = New-AsConnection $port $catalog
  try {
    $tables      = Invoke-Query $conn "SELECT [ID],[Name],[IsHidden] FROM `$SYSTEM.TMSCHEMA_TABLES"
    $columns     = Invoke-Query $conn "SELECT [ID],[TableID],[ExplicitName],[InferredName],[Type],[IsHidden] FROM `$SYSTEM.TMSCHEMA_COLUMNS"
    $measures    = Invoke-Query $conn "SELECT [ID],[TableID],[Name],[IsHidden] FROM `$SYSTEM.TMSCHEMA_MEASURES"
    $hierarchies = Invoke-Query $conn "SELECT [ID],[TableID],[Name],[IsHidden] FROM `$SYSTEM.TMSCHEMA_HIERARCHIES"

    $perspRows   = Invoke-Query $conn "SELECT [ID],[Name] FROM `$SYSTEM.TMSCHEMA_PERSPECTIVES"
    $pTables     = Invoke-Query $conn "SELECT [ID],[PerspectiveID],[TableID] FROM `$SYSTEM.TMSCHEMA_PERSPECTIVE_TABLES"
    $pCols       = Invoke-Query $conn "SELECT [PerspectiveTableID],[ColumnID] FROM `$SYSTEM.TMSCHEMA_PERSPECTIVE_COLUMNS"
    $pMeas       = Invoke-Query $conn "SELECT [PerspectiveTableID],[MeasureID] FROM `$SYSTEM.TMSCHEMA_PERSPECTIVE_MEASURES"
    $pHier       = Invoke-Query $conn "SELECT [PerspectiveTableID],[HierarchyID] FROM `$SYSTEM.TMSCHEMA_PERSPECTIVE_HIERARCHIES"

    $tablesOut = @()
    foreach ($t in $tables | Sort-Object Name) {
      $children = @()
      foreach ($c in $columns | Where-Object { [string]$_.TableID -eq [string]$t.ID }) {
        $name = if ($c.ExplicitName) { $c.ExplicitName } else { $c.InferredName }
        if (-not $name -or $name -like "RowNumber-*") { continue }
        $children += [ordered]@{ id="c$($c.ID)"; name=$name; type='column'; hidden=[bool]$c.IsHidden }
      }
      foreach ($me in $measures | Where-Object { [string]$_.TableID -eq [string]$t.ID }) {
        $children += [ordered]@{ id="m$($me.ID)"; name=$me.Name; type='measure'; hidden=[bool]$me.IsHidden }
      }
      foreach ($h in $hierarchies | Where-Object { [string]$_.TableID -eq [string]$t.ID }) {
        $children += [ordered]@{ id="h$($h.ID)"; name=$h.Name; type='hierarchy'; hidden=[bool]$h.IsHidden }
      }
      $children = @($children | Sort-Object { $_.name })
      $tablesOut += [ordered]@{ id="t$($t.ID)"; name=$t.Name; type='table'; hidden=[bool]$t.IsHidden; expanded=$false; children=@($children) }
    }

    $colName  = @{}; foreach ($c in $columns)     { $n = if ($c.ExplicitName) { $c.ExplicitName } else { $c.InferredName }; $colName[[string]$c.ID] = $n }
    $measName = @{}; foreach ($m in $measures)    { $measName[[string]$m.ID] = $m.Name }
    $hierName = @{}; foreach ($h in $hierarchies) { $hierName[[string]$h.ID] = $h.Name }
    $tblName  = @{}; foreach ($t in $tables)      { $tblName[[string]$t.ID] = $t.Name }

    $perspOut = @()
    foreach ($p in $perspRows) {
      $membership = @()
      $wholeTables = @()
      $myPTs = $pTables | Where-Object { [string]$_.PerspectiveID -eq [string]$p.ID }
      foreach ($pt in $myPTs) {
        $tn = $tblName[[string]$pt.TableID]
        if (-not $tn) { continue }
        $ptCols = @($pCols | Where-Object { [string]$_.PerspectiveTableID -eq [string]$pt.ID })
        $ptMs  = @($pMeas | Where-Object { [string]$_.PerspectiveTableID -eq [string]$pt.ID })
        $ptHs  = @($pHier | Where-Object { [string]$_.PerspectiveTableID -eq [string]$pt.ID })
        $hasAny = ($ptCols.Count + $ptMs.Count + $ptHs.Count) -gt 0
        if (-not $hasAny) {
          $wholeTables += $tn
        } else {
          foreach ($x in $ptCols) { $membership += ($tn + "/" + $colName[[string]$x.ColumnID]) }
          foreach ($x in $ptMs)   { $membership += ($tn + "/" + $measName[[string]$x.MeasureID]) }
          foreach ($x in $ptHs)   { $membership += ($tn + "/" + $hierName[[string]$x.HierarchyID]) }
        }
      }
      $perspOut += [ordered]@{ id = "p$($p.ID)"; name = $p.Name; membership = @($membership); wholeTables = @($wholeTables) }
    }

    return [ordered]@{ catalog = $catalog; tables = @($tablesOut); perspectives = @($perspOut) }
  } finally { $conn.Close() }
}

function Save-Perspectives([int]$port, [object]$body) {
  $server = New-Object Microsoft.AnalysisServices.Tabular.Server
  $server.Connect("Data Source=localhost:$port")
  try {
    $db = $server.Databases[0]
    $model = $db.Model

    if ($body.deletes) {
      foreach ($name in @($body.deletes)) {
        $existing = $model.Perspectives.Find($name)
        if ($existing) { $model.Perspectives.Remove($existing) | Out-Null; Write-Host "Deleted: $name" }
      }
    }

    if ($body.perspectives) {
      foreach ($p in @($body.perspectives)) {
        $existing = $model.Perspectives.Find($p.name)
        if ($existing) { $model.Perspectives.Remove($existing) | Out-Null }

        $newP = New-Object Microsoft.AnalysisServices.Tabular.Perspective
        $newP.Name = $p.name

        foreach ($t in @($p.tables)) {
          $tblObj = $model.Tables.Find($t.name)
          if (-not $tblObj) { Write-Host "  Unknown table: $($t.name)"; continue }

          $pt = New-Object Microsoft.AnalysisServices.Tabular.PerspectiveTable
          $pt.Table = $tblObj

          if (-not $t.wholeTable) {
            if ($t.columns) {
              foreach ($cn in @($t.columns)) {
                $col = $tblObj.Columns.Find($cn)
                if ($col) { $pc = New-Object Microsoft.AnalysisServices.Tabular.PerspectiveColumn; $pc.Column = $col; $pt.PerspectiveColumns.Add($pc) }
              }
            }
            if ($t.measures) {
              foreach ($mn in @($t.measures)) {
                $me = $tblObj.Measures.Find($mn)
                if ($me) { $pm = New-Object Microsoft.AnalysisServices.Tabular.PerspectiveMeasure; $pm.Measure = $me; $pt.PerspectiveMeasures.Add($pm) }
              }
            }
            if ($t.hierarchies) {
              foreach ($hn in @($t.hierarchies)) {
                $h = $tblObj.Hierarchies.Find($hn)
                if ($h) { $ph = New-Object Microsoft.AnalysisServices.Tabular.PerspectiveHierarchy; $ph.Hierarchy = $h; $pt.PerspectiveHierarchies.Add($ph) }
              }
            }
            if ($pt.PerspectiveColumns.Count + $pt.PerspectiveMeasures.Count + $pt.PerspectiveHierarchies.Count -eq 0) { continue }
          }
          $newP.PerspectiveTables.Add($pt)
        }

        $model.Perspectives.Add($newP)
        Write-Host "Saved: $($p.name) ($($newP.PerspectiveTables.Count) tables)"
      }
    }

    $model.SaveChanges() | Out-Null
    return @{ ok = $true; perspectives = @($model.Perspectives | ForEach-Object { $_.Name }) }
  } finally { $server.Disconnect() }
}

# ---- HTTP server ----
$scriptDir = Split-Path -Parent $PSCommandPath
$prefix = "http://localhost:8768/"

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($prefix)
try { $listener.Start() } catch { Write-Error "Bind $prefix failed: $($_.Exception.Message)"; exit 1 }

Write-Host ""
Write-Host "  Perspective Editor proxy v3 (TOM writes) at $prefix" -ForegroundColor Green
Write-Host "  Ctrl+C to stop"
Write-Host ""

function Send-Json($ctx, [int]$status, $obj) {
  $ctx.Response.StatusCode = $status
  $ctx.Response.ContentType = "application/json; charset=utf-8"
  $ctx.Response.Headers["Access-Control-Allow-Origin"] = "*"
  $ctx.Response.Headers["Access-Control-Allow-Methods"] = "GET,POST,OPTIONS"
  $ctx.Response.Headers["Access-Control-Allow-Headers"] = "Content-Type"
  $json = $obj | ConvertTo-Json -Depth 30 -Compress
  $bytes = [Text.Encoding]::UTF8.GetBytes($json)
  $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
  $ctx.Response.OutputStream.Close()
}

function Send-File($ctx, [string]$path, [string]$mime) {
  $ctx.Response.StatusCode = 200
  $ctx.Response.ContentType = $mime
  $ctx.Response.Headers["Access-Control-Allow-Origin"] = "*"
  $bytes = [IO.File]::ReadAllBytes($path)
  $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
  $ctx.Response.OutputStream.Close()
}

while ($listener.IsListening) {
  try { $ctx = $listener.GetContext() } catch { break }
  $req = $ctx.Request
  $path = $req.Url.AbsolutePath
  $method = $req.HttpMethod
  Write-Host "$method $path$($req.Url.Query)"

  try {
    if ($method -eq "OPTIONS") {
      $ctx.Response.Headers["Access-Control-Allow-Origin"] = "*"
      $ctx.Response.Headers["Access-Control-Allow-Methods"] = "GET,POST,OPTIONS"
      $ctx.Response.Headers["Access-Control-Allow-Headers"] = "Content-Type"
      $ctx.Response.StatusCode = 204
      $ctx.Response.OutputStream.Close()
      continue
    }
    if ($path -eq "/" -or $path -eq "/perspective-editor-live.html") {
      $html = Join-Path $scriptDir "perspective-editor-live.html"
      if (Test-Path $html) { Send-File $ctx $html "text/html; charset=utf-8" }
      else { Send-Json $ctx 404 @{ error = "perspective-editor-live.html not found" } }
      continue
    }
    if ($path -eq "/api/model" -and $method -eq "GET") {
      $port = [int]$req.QueryString["port"]
      if (-not $port) { Send-Json $ctx 400 @{ error="port required" }; continue }
      Send-Json $ctx 200 (Get-Model $port)
      continue
    }
    if ($path -eq "/api/perspectives" -and $method -eq "POST") {
      $port = [int]$req.QueryString["port"]
      if (-not $port) { Send-Json $ctx 400 @{ error="port required" }; continue }
      $reader = New-Object IO.StreamReader($req.InputStream, $req.ContentEncoding)
      $body = $reader.ReadToEnd() | ConvertFrom-Json
      Send-Json $ctx 200 (Save-Perspectives $port $body)
      continue
    }
    Send-Json $ctx 404 @{ error = "unknown route $path" }
  } catch {
    Write-Host "ERR: $($_.Exception.Message)" -ForegroundColor Red
    try { Send-Json $ctx 500 @{ error = $_.Exception.Message; type = $_.Exception.GetType().FullName } } catch {}
  }
}
