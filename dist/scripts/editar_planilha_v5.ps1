param(
    [Parameter(Mandatory = $true)][string]$Perfil,
    [Parameter(Mandatory = $true)][string]$Entrada,
    [Parameter(Mandatory = $true)][string]$SaidaDir
)
$ErrorActionPreference = 'Stop'
Write-Host '       (script versão 6)' -ForegroundColor DarkGray
$inv = [System.Globalization.CultureInfo]::InvariantCulture

function Norm([string]$s) {
    $d = ($s + '').Trim().Normalize([Text.NormalizationForm]::FormD)
    return (($d -replace '\p{Mn}', '')).ToUpper()
}
function ParaNumero($v) {
    if ($v -is [double] -or $v -is [int]) { return [double]$v }
    $t = (($v + '') -replace 'R\$', '').Trim()
    if ($t -eq '') { return '' }
    if ($t.Contains(',')) { $t = $t.Replace('.', '').Replace(',', '.') }
    $n = 0.0
    if ([double]::TryParse($t, [Globalization.NumberStyles]::Float, $inv, [ref]$n)) { return $n }
    return $v
}
function GravaCelula($cel, $val) {
    if ($val -is [double]) {
        try { $cel.Value2 = $val; return } catch {}
        try { $cel.Value = $val; return } catch {}
        $cel.Formula = '=' + $val.ToString('R', $inv)
    }
    else {
        if ("$val" -eq '') { return }
        $cel.Value2 = [string]$val
    }
}
function Soltar($o) { if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} } }

$excel = $null; $wbIn = $null; $wb = $null
try {
    $p = Get-Content -Raw -Encoding UTF8 $Perfil | ConvertFrom-Json
    $vb = $p.valor_bruto

    # ------------------------------------------------------------ leitura
    $ext = [IO.Path]::GetExtension($Entrada).ToLower()
    $cab = @(); $linhas = New-Object System.Collections.ArrayList
    if ($ext -eq '.csv' -or $ext -eq '.txt') {
        $cab = @(((Get-Content -Path $Entrada -TotalCount 1 -Encoding UTF8) -replace '^﻿', '') -split ';' | ForEach-Object { $_.Trim().Trim('"') })
        $csv = Import-Csv -Path $Entrada -Delimiter ';' -Encoding UTF8 -Header $cab | Select-Object -Skip 1
        foreach ($r in $csv) {
            $arr = foreach ($c in $cab) { $r.$c }
            [void]$linhas.Add(@($arr))
        }
    }
    else {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false
        $wbIn = $excel.Workbooks.Open($Entrada, 0, $true)
        $v = $wbIn.Worksheets.Item(1).UsedRange.Value2
        $nr = $v.GetUpperBound(0); $nc = $v.GetUpperBound(1)
        $cab = @(1..$nc | ForEach-Object { ($v[1, $_] + '').Trim() })
        for ($i = 2; $i -le $nr; $i++) {
            [void]$linhas.Add(@(1..$nc | ForEach-Object { $v[$i, $_] }))
        }
        $wbIn.Close($false)
    }
    # descarta linhas totalmente vazias
    $linhas = @($linhas | Where-Object { ($_ | Where-Object { "$_".Trim() -ne '' }).Count -gt 0 })
    $nc = $cab.Count; $n = $linhas.Count

    $idx = @{}
    for ($i = 0; $i -lt $nc; $i++) { $k = Norm $cab[$i]; if (-not $idx.ContainsKey($k)) { $idx[$k] = $i + 1 } }

    $necessarias = @($p.colunas_visiveis) + @($p.colunas_numericas) + @($p.coluna_duplicados, $vb.coluna_base, $vb.coluna_uf)
    $faltando = @($necessarias | Sort-Object -Unique | Where-Object { -not $idx.ContainsKey((Norm $_)) })
    if ($faltando.Count -gt 0) {
        Write-Host ('[ERRO] Colunas não encontradas na planilha: ' + ($faltando -join ', ')) -ForegroundColor Red
        Write-Host '       Confira se este é o arquivo certo (cabeçalho na primeira linha).' -ForegroundColor Yellow
        exit 1
    }
    $cNum = @($p.colunas_numericas | ForEach-Object { $idx[(Norm $_)] })
    $cAc = $idx[(Norm $p.coluna_duplicados)]
    $cBase = $idx[(Norm $vb.coluna_base)]
    $cUf = $idx[(Norm $vb.coluna_uf)]
    $cBruto = $nc + 1
    $cPreco = $idx[(Norm 'PRECO_VIGENTE_DO_CIRCUITO')]
    $ultima = $n + 1

    # ------------------------------------------------------------ Excel
    if (-not $excel) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false
    }
    Write-Host ('       (aplicativo: ' + $excel.Name + ' ' + $excel.Version + ')') -ForegroundColor DarkGray
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1); $ws.Name = 'Planilha'
    $wa = $wb.Worksheets.Add([Type]::Missing, $ws); $wa.Name = 'ALIQUOTAS'

    # aba de alíquotas
    $wa.Cells.Item(1, 1).Value2 = 'UF'; $wa.Cells.Item(1, 2).Value2 = 'ICMS'
    $r = 2
    foreach ($prop in $p.aliquotas_icms.PSObject.Properties) {
        $wa.Cells.Item($r, 1).Value2 = $prop.Name
        GravaCelula $wa.Cells.Item($r, 2) ([double]$prop.Value)
        $r++
    }
    $wa.Range("B2:B$($r-1)").NumberFormat = '0.0%'
    $wa.Cells.Item(1, 4).Value2 = 'PIS/COFINS'
    GravaCelula $wa.Cells.Item(1, 5) ([double]$vb.pis_cofins)
    $wa.Cells.Item(1, 5).NumberFormat = '0.00%'
    $wa.Range('A1:B1').Font.Bold = $true; $wa.Range('D1').Font.Bold = $true

    # dados (tudo como texto, exceto colunas numéricas)
    $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($ultima, $nc)).NumberFormat = '@'
    foreach ($c in $cNum) { $ws.Range($ws.Cells.Item(2, $c), $ws.Cells.Item($ultima, $c)).NumberFormat = '#,##0.00' }
    # escreve célula a célula (compatível com Excel e similares)
    for ($i = 0; $i -le $n; $i++) {
        for ($j = 0; $j -lt $nc; $j++) {
            if ($i -eq 0) { $val = [string]$cab[$j] }
            else {
                $val = $linhas[$i - 1][$j]
                if ($cNum -contains ($j + 1)) { $val = ParaNumero $val }
                if (($j + 1) -eq $cPreco -and $val -is [double]) { $val = [Math]::Round($val, 2) }
            }
            GravaCelula $ws.Cells.Item($i + 1, $j + 1) $val
        }
    }

    # VALOR_BRUTO
    $ws.Cells.Item(1, $cBruto).Value2 = $vb.nome
    if ($n -ge 1) {
        $LB = $ws.Cells.Item(1, $cBase).Address($false, $false) -replace '\d', ''
        $LU = $ws.Cells.Item(1, $cUf).Address($false, $false) -replace '\d', ''
        $f = '=IF(' + $LB + '2="","",IFERROR(ROUND(' + $LB + '2/((1-VLOOKUP(' + $LU + '2,ALIQUOTAS!$A:$B,2,FALSE))*(1-ALIQUOTAS!$E$1)),2),""))'
        $rg = $ws.Range($ws.Cells.Item(2, $cBruto), $ws.Cells.Item($ultima, $cBruto))
        $rg.Formula = $f
        $rg.NumberFormat = '#,##0.00'
        $ws.Range($ws.Cells.Item(2, $cPreco), $ws.Cells.Item($ultima, $cPreco)).NumberFormat = '#,##0.00'

        # duplicados em ACESSO -> vermelho
        $rd = $ws.Range($ws.Cells.Item(2, $cAc), $ws.Cells.Item($ultima, $cAc))
        $LA = $ws.Cells.Item(1, $cAc).Address($false, $false) -replace '\d', ''
        try {
            $fc = $rd.FormatConditions.AddUniqueValues()
            $fc.DupeUnique = 1            # xlDuplicate
        }
        catch {
            $fc = $rd.FormatConditions.Add(2, [Type]::Missing, ('=AND(' + $LA + '2<>"",COUNTIF($' + $LA + '$2:$' + $LA + '$' + $ultima + ',' + $LA + '2)>1)'))
        }
        $fc.Interior.Color = 255      # vermelho
        try { $fc.Font.Color = 16777215; $fc.Font.Bold = $true } catch {}
    }

    # visibilidade e aparência
    $visiveis = @($p.colunas_visiveis | ForEach-Object { $idx[(Norm $_)] }) + $cBruto
    for ($c = 1; $c -le $cBruto; $c++) {
        $col = $ws.Columns.Item($c)
        if ($visiveis -contains $c) { $col.ColumnWidth = 26 } else { $col.Hidden = $true }
    }
    $hd = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(1, $cBruto))
    $hd.Font.Bold = $true; $hd.Interior.Color = 15917529
    $hd.HorizontalAlignment = -4108
    try { [void]$ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($ultima, $cBruto)).AutoFilter() } catch {}
    try { $ws.Activate(); $excel.ActiveWindow.SplitRow = 1; $excel.ActiveWindow.FreezePanes = $true } catch {}

    # salvar
    if (-not (Test-Path $SaidaDir)) { New-Item -ItemType Directory -Path $SaidaDir | Out-Null }
    $base = [IO.Path]::GetFileNameWithoutExtension($Entrada)
    $dest = Join-Path $SaidaDir ($base + '_EDITADA.xlsx')
    $wb.SaveAs($dest, 51)
    $wb.Close($true)
    Write-Host "[OK] $n linha(s) processada(s)." -ForegroundColor Green
    Write-Host "[OK] Arquivo salvo em: $dest" -ForegroundColor Green
    exit 0
}
catch {
    $m = $_.Exception.Message
    if ($m -match 'Excel.Application|0x80040154|REGDB') {
        Write-Host '[ERRO] Microsoft Excel não encontrado neste computador.' -ForegroundColor Red
    }
    elseif ($m -match 'being used|em uso|used by another') {
        Write-Host '[ERRO] Não consegui salvar. Feche a planilha _EDITADA se ela estiver aberta no Excel.' -ForegroundColor Red
    }
    else { Write-Host "[ERRO] $m" -ForegroundColor Red }
    Write-Host ("       (linha " + $_.InvocationInfo.ScriptLineNumber + " do script: " + $_.InvocationInfo.Line.Trim() + ")") -ForegroundColor DarkGray
    exit 1
}
finally {
    if ($excel) { try { $excel.Quit() } catch {} }
    Soltar $wb; Soltar $wbIn; Soltar $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
