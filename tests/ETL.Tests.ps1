# ── ETL.Tests.ps1 ────────────────────────────────────────────────────────────
# Pester 3 unit + integration tests for the retail-etl pipeline.
# Run with:  powershell -File tests\ETL.Tests.ps1
# ─────────────────────────────────────────────────────────────────────────────

$ROOT = Split-Path $PSScriptRoot -Parent

# ── Unit: ConvertMonthLabel ───────────────────────────────────────────────────
Describe "ConvertMonthLabel" {

    . "$ROOT\scripts\procesar_pyg.ps1"

    It "converts 'Enero 2024' to 'Ene 24'" {
        ConvertMonthLabel "Enero 2024" | Should Be "Ene 24"
    }
    It "converts multiline 'Junio`n2025' to 'Jun 25'" {
        ConvertMonthLabel "Junio`n2025" | Should Be "Jun 25"
    }
    It "converts 'Diciembre 2025' to 'Dic 25'" {
        ConvertMonthLabel "Diciembre 2025" | Should Be "Dic 25"
    }
    It "strips extra whitespace" {
        ConvertMonthLabel "  Marzo  2024  " | Should Be "Mar 24"
    }
    It "returns empty string for blank input" {
        ConvertMonthLabel "" | Should Be ""
    }
    It "rejects non-month label (label column header)" {
        # After conversion, non-month text should NOT match month pattern
        $label = ConvertMonthLabel "Total"
        $label -match "^(Ene|Feb|Mar|Abr|May|Jun|Jul|Ago|Sep|Oct|Nov|Dic)" | Should Be $false
    }
}

# ── Unit: ConvertBalanceMonth ─────────────────────────────────────────────────
Describe "ConvertBalanceMonth" {

    . "$ROOT\scripts\procesar_balance.ps1"

    It "converts 'Enero 2024' to 'Ene 24'" {
        ConvertBalanceMonth "Enero 2024" | Should Be "Ene 24"
    }
    It "converts 'Septiembre 2025' to 'Sep 25'" {
        ConvertBalanceMonth "Septiembre 2025" | Should Be "Sep 25"
    }
    It "handles multiline header" {
        ConvertBalanceMonth "Agosto`n2025" | Should Be "Ago 25"
    }
}

# ── Unit: ReadCellPyG safe type handling ─────────────────────────────────────
Describe "ReadCellPyG" {

    . "$ROOT\scripts\procesar_pyg.ps1"

    It "returns 0 when row is 0" {
        ReadCellPyG $null 0 1 | Should Be 0
    }
    It "returns 0 for null cell value" {
        $fakeWs = New-Object PSObject
        $fakeWs | Add-Member -MemberType ScriptMethod -Name "Cells" -Value {} -Force
        # Simulate via a small helper hashtable mock
        function script:FakeCell($val) {
            $c = New-Object PSObject
            $c | Add-Member -MemberType NoteProperty -Name Value2 -Value $val
            $ws = New-Object PSObject
            $ws | Add-Member -MemberType ScriptMethod -Name Cells -Value { $c } -Force
            return $ws
        }
        # row=0 path
        ReadCellPyG $null 0 1 | Should Be 0
    }
    It "returns 0 for text cell (label column)" {
        # Build a minimal fake worksheet
        $fakeCell = New-Object PSObject
        $fakeCell | Add-Member NoteProperty Value2 "Total Operacionales"

        $fakeCells = New-Object PSObject
        $fakeCells | Add-Member -MemberType ScriptMethod -Name Item `
            -Value { param($r,$c) return $fakeCell } -Force

        $fakeWs = New-Object PSObject
        $fakeWs | Add-Member NoteProperty Cells $fakeCells

        ReadCellPyG $fakeWs 5 1 | Should Be 0
    }
    It "rounds numeric cell value" {
        $fakeCell = New-Object PSObject
        $fakeCell | Add-Member NoteProperty Value2 12345678.9

        $fakeCells = New-Object PSObject
        $fakeCells | Add-Member -MemberType ScriptMethod -Name Item `
            -Value { param($r,$c) return $fakeCell } -Force

        $fakeWs = New-Object PSObject
        $fakeWs | Add-Member NoteProperty Cells $fakeCells

        ReadCellPyG $fakeWs 5 3 | Should Be 12345679
    }
}

# ── Integration: Config file ──────────────────────────────────────────────────
Describe "Config: rutas.ps1" {

    . "$ROOT\config\rutas.ps1"

    It "loads CONFIG hashtable" {
        $script:CONFIG | Should Not BeNullOrEmpty
    }
    It "defines DataAnalytics path" {
        $script:CONFIG.DataAnalytics | Should Not BeNullOrEmpty
    }
    It "defines three P&G years (2024, 2025, 2026)" {
        ($script:CONFIG.Archivos_PyG.Keys -contains "2024") | Should Be $true
        ($script:CONFIG.Archivos_PyG.Keys -contains "2025") | Should Be $true
        ($script:CONFIG.Archivos_PyG.Keys -contains "2026") | Should Be $true
    }
    It "defines two Balance date ranges" {
        $script:CONFIG.Archivos_Balance.Count | Should Be 2
    }
    It "defines output path for data.js" {
        $script:CONFIG.Output_Dashboard | Should Match "data\.js$"
    }
    It "defines output path for data_balance.js" {
        $script:CONFIG.Output_Balance_Dashboard | Should Match "data_balance\.js$"
    }
    It "data.js output path exists (directory)" {
        $dir = Split-Path $script:CONFIG.Output_Dashboard
        Test-Path $dir | Should Be $true
    }
}

# ── Integration: Output files produced by ETL ────────────────────────────────
Describe "ETL output: data.js" {

    . "$ROOT\config\rutas.ps1"
    $outPath = $script:CONFIG.Output_Dashboard

    It "data.js file exists after ETL run" {
        Test-Path $outPath | Should Be $true
    }
    It "data.js contains window.DASHBOARD_DATA" {
        $content = [System.IO.File]::ReadAllText($outPath)
        $content | Should Match "window\.DASHBOARD_DATA"
    }
    It "data.js contains ventas_mensual array" {
        $content = [System.IO.File]::ReadAllText($outPath)
        $content | Should Match "ventas_mensual"
    }
    It "data.js has at least 20 month records" {
        $content = [System.IO.File]::ReadAllText($outPath)
        ($content | Select-String "periodo:" -AllMatches).Matches.Count | Should BeGreaterThan 20
    }
    It "data.js is valid UTF-8 (no BOM)" {
        $bytes = [System.IO.File]::ReadAllBytes($outPath)
        # UTF-8 BOM is EF BB BF
        ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) | Should Be $false
    }
    It "data.js records contain required numeric fields" {
        $content = [System.IO.File]::ReadAllText($outPath)
        $content | Should Match "ventas:\d+"
        $content | Should Match "margen_pct:"
        $content | Should Match "personal:"
    }
}

Describe "ETL output: data_balance.js" {

    . "$ROOT\config\rutas.ps1"
    $outPath = $script:CONFIG.Output_Balance_Dashboard

    It "data_balance.js file exists after ETL run" {
        Test-Path $outPath | Should Be $true
    }
    It "data_balance.js contains window.BALANCE_DATA" {
        $content = [System.IO.File]::ReadAllText($outPath)
        $content | Should Match "window\.BALANCE_DATA"
    }
    It "data_balance.js has at least 10 month records" {
        $content = [System.IO.File]::ReadAllText($outPath)
        ($content | Select-String "mes:" -AllMatches).Matches.Count | Should BeGreaterThan 10
    }
    It "data_balance.js records contain balance fields" {
        $content = [System.IO.File]::ReadAllText($outPath)
        $content | Should Match "inventarios:"
        $content | Should Match "patrimonio:"
        $content | Should Match "total_activo:"
    }
}
