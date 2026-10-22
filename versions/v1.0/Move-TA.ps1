<#
.SYNOPSIS
    Move TA - realocacao de Technical Accounts no STAR IPS 3.62.x (GUI).

.DESCRIPTION
    Interface WinForms por cima das procedures dbo.usp_IPS_MoveTA / dbo.usp_IPS_RetireTA
    (arquivo usp_IPS_MoveTA.sql, na mesma pasta; a tool oferece instalar se faltarem).

    Fluxo: Conectar -> escolher planta -> marcar TAs na arvore (subaccounts entram junto)
           -> "Definir destino..." pergunta: trocar prefixo (801. -> 701.) ou escolher
              o codigo de cada TA -> Simular (tudo numa transacao + ROLLBACK)
           -> Executar (mesma transacao + COMMIT).

.PARAMETER SelfTest
    Roda sem interface: conecta, monta o mapeamento de teste e simula (rollback). Uso
    interno para validar o script.
#>
[CmdletBinding()]
param(
    [string]$Server   = 'localhost\SQL2025',
    [string]$Database = 'ADEF',
    [switch]$SelfTest,
    [switch]$UiTest,
    [string]$ScreenshotDir,
    [string]$TestPlant = 'FMHUB',
    [string[]]$TestSources = @('801.01','801.04','801.10'),
    [string]$TestOldPrefix = '801.',
    [string]$TestNewPrefix = '701.'
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data

$script:ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:ProcFile  = Join-Path $script:ScriptDir 'usp_IPS_MoveTA.sql'

# ----------------------------------------------------------------------------- estado
$script:Conn      = $null          # SqlConnection aberta
$script:FuncClass = $null
$script:Tree      = @{}            # Function_Id -> @{Id;Father;Name;Obj;InPlant}
$script:ClassSet  = @{}            # Function_Id -> 1 (ativa na classe, mesmo sem estar na planta)
$script:Links     = @{}            # Function_Id -> texto com contagem de ligacoes
$script:Mapping   = @()            # lista de @{From;FromName;To;ToName;Status;Links;IsTop}
$script:Checking  = $false

# ----------------------------------------------------------------------------- SQL helpers
function New-ConnString {
    param([string]$Srv, [string]$Db, [bool]$WinAuth, [string]$User, [string]$Pass)
    if ($WinAuth) { return "Server=$Srv;Database=$Db;Integrated Security=True;TrustServerCertificate=True;Application Name=MoveTA" }
    return "Server=$Srv;Database=$Db;User ID=$User;Password=$Pass;TrustServerCertificate=True;Application Name=MoveTA"
}
function Invoke-Sql {
    param([string]$Sql, [hashtable]$Params = @{}, [System.Data.SqlClient.SqlTransaction]$Tran = $null, [int]$Timeout = 600)
    $cmd = $script:Conn.CreateCommand()
    $cmd.CommandText = $Sql
    $cmd.CommandTimeout = $Timeout
    if ($Tran) { $cmd.Transaction = $Tran }
    foreach ($k in $Params.Keys) { $v = $Params[$k]; if ($null -eq $v) { $v = [DBNull]::Value }; [void]$cmd.Parameters.AddWithValue($k, $v) }
    $tables = New-Object System.Collections.ArrayList
    $rd = $cmd.ExecuteReader()
    try {
        do { $t = New-Object System.Data.DataTable; $t.Load($rd); [void]$tables.Add($t) } while (-not $rd.IsClosed -and $rd.HasRows -or ($rd.FieldCount -gt 0 -and -not $rd.IsClosed))
    } finally { $rd.Close() }
    return ,$tables
}
function Invoke-SqlScalar { param([string]$Sql, [hashtable]$Params = @{})
    $cmd = $script:Conn.CreateCommand(); $cmd.CommandText = $Sql
    foreach ($k in $Params.Keys) { [void]$cmd.Parameters.AddWithValue($k, $Params[$k]) }
    return $cmd.ExecuteScalar()
}
function Test-Procs { $r = Invoke-SqlScalar "SELECT CASE WHEN OBJECT_ID('dbo.usp_IPS_MoveTA') IS NOT NULL AND OBJECT_ID('dbo.usp_IPS_RetireTA') IS NOT NULL THEN 1 ELSE 0 END"; return ($r -eq 1) }
function Install-Procs {
    if (-not (Test-Path $script:ProcFile)) { throw "Arquivo $($script:ProcFile) nao encontrado." }
    $text = Get-Content $script:ProcFile -Raw
    $batches = [regex]::Split($text, '(?im)^\s*GO\s*$')
    foreach ($b in $batches) { if ($b.Trim().Length -gt 0) { $cmd = $script:Conn.CreateCommand(); $cmd.CommandText = $b; $cmd.CommandTimeout = 120; [void]$cmd.ExecuteNonQuery() } }
}

# ----------------------------------------------------------------------------- dados
function Load-Plants {
    $t = (Invoke-Sql "SELECT Plant_Id, Plant_Name, FuncClass_Id FROM ST_Plant WHERE Rec_Deleted = 0 ORDER BY Plant_Name")[0]
    return ,$t   # virgula: evita que o PowerShell desenrole a DataTable em linhas
}
function Load-Tree { param([string]$Plant)
    $script:FuncClass = Invoke-SqlScalar "SELECT FuncClass_Id FROM ST_Plant WHERE Plant_Id = @p" @{ '@p' = $Plant }
    $sql = @"
SELECT f.Function_Id, f.Function_Father, f.Function_Name, pf.PLANTFUNC_OBJECT_ID,
       CASE WHEN pf.Function_Id IS NULL OR f.Rec_Deleted <> 0 THEN 0 ELSE 1 END AS InPlant,
       CASE WHEN f.Rec_Deleted = 0 THEN 1 ELSE 0 END AS Active
  FROM ST_Function f
  LEFT JOIN ST_PlantFunction pf ON pf.FuncClass_Id = f.FuncClass_Id AND pf.Function_Id = f.Function_Id AND pf.Plant_Id = @p AND pf.Rec_Deleted = 0
 WHERE f.FuncClass_Id = @fc
 ORDER BY f.Function_Id
"@
    $t = (Invoke-Sql $sql @{ '@p' = $Plant; '@fc' = $script:FuncClass })[0]
    $script:Tree = @{}; $script:ClassSet = @{}; $script:DeletedSet = @{}
    foreach ($r in $t.Rows) {
        $id = [string]$r.Function_Id
        if ([int]$r.Active -eq 0) { $script:DeletedSet[$id] = 1; continue }
        $script:ClassSet[$id] = 1
        if ([int]$r.InPlant -eq 1) {
            $script:Tree[$id] = @{ Id = $id; Father = $(if ($r.Function_Father -is [DBNull]) { $null } else { [string]$r.Function_Father }); Name = [string]$r.Function_Name; Obj = $r.PLANTFUNC_OBJECT_ID }
        }
    }
    # contagem de ligacoes por TA (para a coluna "Ligacoes")
    $sql2 = @"
SELECT x.Function_Id, x.Kind, COUNT(*) n FROM (
    SELECT pfo.Function_Id, CASE pfo.ObjType_Id WHEN 320 THEN 'jobs' WHEN 340 THEN 'WOs' WHEN 221 THEN 'spares' WHEN 210 THEN 'counters' WHEN 250 THEN 'equip' WHEN 285 THEN 'assemblies' ELSE 'outros' END Kind
      FROM ST_PlantFunctionObject pfo WHERE pfo.Plant_Id = @p AND pfo.FuncClass_Id = @fc AND pfo.Rec_Deleted = 0
    UNION ALL
    SELECT wh.WH_Function_Id, 'WH' FROM ST_WorkHistory wh WHERE wh.Plant_Id = @p AND wh.WH_FuncClass_Id = @fc AND wh.Rec_Deleted = 0
    UNION ALL
    SELECT pf.Function_Id, 'where-list' FROM ST_WPWHERE2 w JOIN ST_PlantFunction pf ON pf.PLANTFUNC_OBJECT_ID = w.OBJECT_ID WHERE pf.Plant_Id = @p AND w.REC_DELETED = 0 AND ISNULL(w.WPWHERE_MAIN,0) <> -1
) x GROUP BY x.Function_Id, x.Kind
"@
    $t2 = (Invoke-Sql $sql2 @{ '@p' = $Plant; '@fc' = $script:FuncClass })[0]
    $script:Links = @{}
    foreach ($r in $t2.Rows) { $id = [string]$r.Function_Id; $s = "$($r.Kind) $($r.n)"; if ($script:Links.ContainsKey($id)) { $script:Links[$id] += ", $s" } else { $script:Links[$id] = $s } }
}
function Get-Descendants { param([string]$Id)
    $out = New-Object System.Collections.ArrayList
    foreach ($k in $script:Tree.Keys) { if ($script:Tree[$k].Father -eq $Id) { [void]$out.Add($k); foreach ($d in (Get-Descendants $k)) { [void]$out.Add($d) } } }
    return ,$out
}
function Get-TopLevel { param([string[]]$Ids)
    $set = @{}; foreach ($i in $Ids) { $set[$i] = 1 }
    $top = @()
    foreach ($i in $Ids) {
        $f = $script:Tree[$i].Father; $isTop = $true
        while ($f) { if ($set.ContainsKey($f)) { $isTop = $false; break }; if ($script:Tree.ContainsKey($f)) { $f = $script:Tree[$f].Father } else { $f = $null } }
        if ($isTop) { $top += $i }
    }
    return ,$top
}
function Build-Mapping {
    <# $Selected: ids marcados (ja incluindo descendentes). $TopMap: hashtable origem-top -> @{To;ToName} #>
    param([string[]]$Selected, [hashtable]$TopMap)
    $set = @{}; foreach ($i in $Selected) { $set[$i] = 1 }
    $rows = @()
    foreach ($id in ($Selected | Sort-Object { $_.Length }, { $_ })) {
        $to = $null; $toName = $script:Tree[$id].Name; $isTop = $false
        if ($TopMap.ContainsKey($id)) { $to = $TopMap[$id].To; if ($TopMap[$id].ToName) { $toName = $TopMap[$id].ToName }; $isTop = $true }
        else {
            # acha o ancestral top mapeado e troca o prefixo
            $anc = $script:Tree[$id].Father
            while ($anc -and -not $TopMap.ContainsKey($anc)) { if ($script:Tree.ContainsKey($anc)) { $anc = $script:Tree[$anc].Father } else { $anc = $null } }
            if ($anc -and $id.StartsWith($anc)) { $to = $TopMap[$anc].To + $id.Substring($anc.Length) }
        }
        $status = 'INVALIDO (sem destino)'
        if ($to) {
            if ($to -eq $id) { $status = 'INVALIDO (igual a origem)' }
            elseif ($script:Tree.ContainsKey($to)) { $status = 'MERGE - ja existe na planta: ' + $script:Tree[$to].Name }
            elseif ($script:ClassSet.ContainsKey($to)) { $status = 'existe na classe, sera ativada na planta' }
            elseif ($script:DeletedSet.ContainsKey($to)) { $status = 'existia desativada, sera reativada' }
            else { $status = 'sera criada' }
        }
        $links = $(if ($script:Links.ContainsKey($id)) { $script:Links[$id] } else { '-' })
        $rows += @{ From = $id; FromName = $script:Tree[$id].Name; To = $to; ToName = $toName; Status = $status; Links = $links; IsTop = $isTop }
    }
    return ,$rows
}
function Run-Moves {
    <# Executa todos os pares numa unica transacao. $Commit=$false => ROLLBACK (simulacao). Devolve DataTable de log. #>
    param([array]$Map, [string]$Plant, [bool]$Commit, [hashtable]$Opt)
    $log = New-Object System.Data.DataTable
    foreach ($c in 'Par','Tabela','Acao','Linhas','Detalhe') { [void]$log.Columns.Add($c) }
    $tran = $script:Conn.BeginTransaction('MoveTA')
    try {
        foreach ($m in $Map) {
            if ($m.Status -like 'INVALIDO*') { throw "Par $($m.From) invalido: $($m.Status)" }
            $p = @{ '@Plant_Id' = $Plant; '@From_Function_Id' = $m.From; '@To_Function_Id' = $m.To; '@To_Function_Name' = $m.ToName
                    '@MoveWorkHistory' = [int]$Opt.WH; '@MovePOItems' = [int]$Opt.PO; '@MoveMaterialRefs' = [int]$Opt.Mat; '@MoveOtherRefs' = [int]$Opt.Other
                    '@RetireSource' = 0; '@ListItems' = [int]$Opt.ListItems; '@DryRun' = 0 }
            $res = Invoke-Sql "EXEC dbo.usp_IPS_MoveTA @Plant_Id, @From_Function_Id, @To_Function_Id, @To_Function_Name, 1, @MoveWorkHistory, @MovePOItems, @MoveMaterialRefs, @MoveOtherRefs, @RetireSource, 0, @ListItems, NULL, @DryRun" $p $tran
            foreach ($r in $res[0].Rows) { [void]$log.Rows.Add("$($m.From) -> $($m.To)", $r.Tabela, $r.Acao, $r.Linhas, $r.Detalhe) }
        }
        if ($Opt.Retire) {
            # filhos primeiro (codigo mais longo), depois os pais
            foreach ($m in ($Map | Sort-Object { -$_.From.Length }, { $_.From })) {
                $res = Invoke-Sql "EXEC dbo.usp_IPS_RetireTA @Plant_Id, @Function_Id, @RetireInClass, NULL" @{ '@Plant_Id' = $Plant; '@Function_Id' = $m.From; '@RetireInClass' = [int]$Opt.RetireClass } $tran
                foreach ($r in $res[0].Rows) { [void]$log.Rows.Add("$($m.From) (desativar)", $r.Tabela, $r.Acao, $r.Linhas, $r.Detalhe) }
            }
        }
        if ($Commit) { $tran.Commit(); [void]$log.Rows.Add('*', '-', 'COMMIT', $null, 'Alteracoes gravadas.') }
        else { $tran.Rollback(); [void]$log.Rows.Add('*', '-', 'ROLLBACK', $null, 'SIMULACAO: nada foi gravado.') }
    } catch {
        try { $tran.Rollback() } catch {}
        [void]$log.Rows.Add('*', '-', 'ERRO', $null, $_.Exception.Message + '  (transacao desfeita, nada gravado)')
        $script:LastError = $_.Exception.Message
    }
    return ,$log
}

# ============================================================================= SELF TEST
if ($SelfTest) {
    Write-Host "SelfTest: $Server / $Database / planta $TestPlant"
    $script:Conn = New-Object System.Data.SqlClient.SqlConnection (New-ConnString $Server $Database $true '' '')
    $script:Conn.Open()
    if (-not (Test-Procs)) { Write-Host 'instalando procs...'; Install-Procs }
    Load-Tree $TestPlant
    Write-Host "TAs na planta: $($script:Tree.Count); na classe: $($script:ClassSet.Count)"
    $sel = @(); foreach ($s in $TestSources) { $sel += $s; $sel += (Get-Descendants $s) }
    $top = Get-TopLevel $sel
    Write-Host "selecionadas: $($sel.Count)  top-level: $($top -join ', ')"
    $topMap = @{}; foreach ($t in $top) { if ($t.StartsWith($TestOldPrefix)) { $topMap[$t] = @{ To = $TestNewPrefix + $t.Substring($TestOldPrefix.Length); ToName = $null } } }
    $map = Build-Mapping $sel $topMap
    $map | ForEach-Object { '{0,-14} -> {1,-14} {2,-45} {3}' -f $_.From, $_.To, $_.Status, $_.Links } | Write-Host
    $opt = @{ WH = $true; PO = $true; Mat = $true; Other = $true; Retire = $true; RetireClass = $true; ListItems = $false }
    $log = Run-Moves $map $TestPlant $false $opt
    Write-Host "log rows: $($log.Rows.Count)"
    $log | Where-Object { $_.Acao -in 'criada','reativada','desativada','MANTIDA','ATENCAO','VERIFICAR','ERRO','ROLLBACK','COMMIT' -or ($_.Acao -eq 'movidas' -and [int]$_.Linhas -gt 0) } |
        ForEach-Object { '{0,-28} {1,-24} {2,-12} {3,5}  {4}' -f $_.Par, $_.Tabela, $_.Acao, $_.Linhas, $_.Detalhe } | Write-Host
    $script:Conn.Close()
    exit 0
}

# ============================================================================= GUI
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Move TA - STAR IPS 3.62 - realocacao de Technical Accounts'
$form.Size = New-Object System.Drawing.Size(1250, 820)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

function New-Label { param($Text, $X, $Y, $W = 90) $l = New-Object System.Windows.Forms.Label; $l.Text = $Text; $l.Location = New-Object System.Drawing.Point($X, $Y); $l.Size = New-Object System.Drawing.Size($W, 20); $l }
function New-TextBox { param($X, $Y, $W, $Text = '') $t = New-Object System.Windows.Forms.TextBox; $t.Location = New-Object System.Drawing.Point($X, $Y); $t.Size = New-Object System.Drawing.Size($W, 23); $t.Text = $Text; $t }
function New-Button { param($Text, $X, $Y, $W = 110, $H = 28) $b = New-Object System.Windows.Forms.Button; $b.Text = $Text; $b.Location = New-Object System.Drawing.Point($X, $Y); $b.Size = New-Object System.Drawing.Size($W, $H); $b }
function New-Check { param($Text, $X, $Y, $W = 260, $Checked = $true) $c = New-Object System.Windows.Forms.CheckBox; $c.Text = $Text; $c.Location = New-Object System.Drawing.Point($X, $Y); $c.Size = New-Object System.Drawing.Size($W, 22); $c.Checked = $Checked; $c }

# --- conexao
$gbConn = New-Object System.Windows.Forms.GroupBox; $gbConn.Text = '1. Conexao'; $gbConn.Location = New-Object System.Drawing.Point(10, 8); $gbConn.Size = New-Object System.Drawing.Size(1210, 60)
$gbConn.Controls.Add((New-Label 'Servidor:' 10 25 60)); $tbServer = New-TextBox 70 22 170 $Server; $gbConn.Controls.Add($tbServer)
$gbConn.Controls.Add((New-Label 'Banco:' 250 25 50));   $tbDb = New-TextBox 300 22 120 $Database; $gbConn.Controls.Add($tbDb)
$rbWin = New-Object System.Windows.Forms.RadioButton; $rbWin.Text = 'Windows'; $rbWin.Checked = $true; $rbWin.Location = New-Object System.Drawing.Point(435, 23); $rbWin.Size = New-Object System.Drawing.Size(80, 22); $gbConn.Controls.Add($rbWin)
$rbSql = New-Object System.Windows.Forms.RadioButton; $rbSql.Text = 'SQL login'; $rbSql.Location = New-Object System.Drawing.Point(515, 23); $rbSql.Size = New-Object System.Drawing.Size(85, 22); $gbConn.Controls.Add($rbSql)
$gbConn.Controls.Add((New-Label 'Usuario:' 605 25 55)); $tbUser = New-TextBox 660 22 110; $tbUser.Enabled = $false; $gbConn.Controls.Add($tbUser)
$gbConn.Controls.Add((New-Label 'Senha:' 780 25 45));   $tbPass = New-TextBox 825 22 110; $tbPass.UseSystemPasswordChar = $true; $tbPass.Enabled = $false; $gbConn.Controls.Add($tbPass)
$btnConnect = New-Button 'Conectar' 950 20 100; $gbConn.Controls.Add($btnConnect)
$gbConn.Controls.Add((New-Label 'Planta:' 1060 25 50)); $cbPlant = New-Object System.Windows.Forms.ComboBox; $cbPlant.Location = New-Object System.Drawing.Point(1105, 22); $cbPlant.Size = New-Object System.Drawing.Size(95, 23); $cbPlant.DropDownStyle = 'DropDownList'; $cbPlant.Enabled = $false; $gbConn.Controls.Add($cbPlant)
$form.Controls.Add($gbConn)
$rbSql.add_CheckedChanged({ $tbUser.Enabled = $rbSql.Checked; $tbPass.Enabled = $rbSql.Checked })

# --- arvore
$gbTree = New-Object System.Windows.Forms.GroupBox; $gbTree.Text = '2. TAs de origem (marque; os subaccounts entram junto)'; $gbTree.Location = New-Object System.Drawing.Point(10, 75); $gbTree.Size = New-Object System.Drawing.Size(470, 640)
$gbTree.Controls.Add((New-Label 'Filtro:' 10 25 45)); $tbFilter = New-TextBox 55 22 300; $gbTree.Controls.Add($tbFilter)
$btnFilter = New-Button 'Localizar' 360 20 95 25; $gbTree.Controls.Add($btnFilter)
$tv = New-Object System.Windows.Forms.TreeView; $tv.Location = New-Object System.Drawing.Point(10, 55); $tv.Size = New-Object System.Drawing.Size(450, 545); $tv.CheckBoxes = $true; $tv.HideSelection = $false; $gbTree.Controls.Add($tv)
$chkKids = New-Check 'Incluir subaccounts automaticamente' 10 605 300 $true; $gbTree.Controls.Add($chkKids)
$lblSel = New-Label '0 selecionadas' 320 607 140; $gbTree.Controls.Add($lblSel)
$form.Controls.Add($gbTree)

# --- destino / mapeamento
$gbMap = New-Object System.Windows.Forms.GroupBox; $gbMap.Text = '3. Destino'; $gbMap.Location = New-Object System.Drawing.Point(490, 75); $gbMap.Size = New-Object System.Drawing.Size(730, 330)
$btnDest = New-Button 'Definir destino...' 10 22 140; $btnDest.Enabled = $false; $gbMap.Controls.Add($btnDest)
$lblMode = New-Label '' 160 27 550; $gbMap.Controls.Add($lblMode)
$dgMap = New-Object System.Windows.Forms.DataGridView; $dgMap.Location = New-Object System.Drawing.Point(10, 55); $dgMap.Size = New-Object System.Drawing.Size(710, 265)
$dgMap.ReadOnly = $true; $dgMap.AllowUserToAddRows = $false; $dgMap.RowHeadersVisible = $false; $dgMap.AutoSizeColumnsMode = 'Fill'; $dgMap.SelectionMode = 'FullRowSelect'
foreach ($c in @(@('Origem',90),@('Nome',160),@('Destino',90),@('Novo nome',160),@('Situacao',150),@('Ligacoes',110))) { $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $col.HeaderText = $c[0]; $col.Name = $c[0]; $col.FillWeight = $c[1]; [void]$dgMap.Columns.Add($col) }
$gbMap.Controls.Add($dgMap)
$form.Controls.Add($gbMap)

# --- opcoes + acoes
$gbOpt = New-Object System.Windows.Forms.GroupBox; $gbOpt.Text = '4. Opcoes e execucao'; $gbOpt.Location = New-Object System.Drawing.Point(490, 410); $gbOpt.Size = New-Object System.Drawing.Size(730, 95)
$chkWH    = New-Check 'Mover work history (ST_WorkHistory / ST_WHWHERE2)' 10 20 330; $gbOpt.Controls.Add($chkWH)
$chkPO    = New-Check 'Mover linhas de MR/PO (ST_POAITEM2)' 10 42 330; $gbOpt.Controls.Add($chkPO)
$chkMat   = New-Check 'Mover referencias de materiais (default TA, transacoes)' 10 64 330; $gbOpt.Controls.Add($chkMat)
$chkOther = New-Check 'Mover orcamento / custos / restricoes de grupo' 350 20 300; $gbOpt.Controls.Add($chkOther)
$chkRetire = New-Check 'Desativar TAs de origem depois de mover' 350 42 300 $true; $gbOpt.Controls.Add($chkRetire)
$chkRetireClass = New-Check 'Desativar tambem na classe (se nenhuma outra planta usar)' 350 64 370 $true; $gbOpt.Controls.Add($chkRetireClass)
$form.Controls.Add($gbOpt)

$btnSim  = New-Button 'Simular (sem gravar)' 490 512 160 32; $btnSim.Enabled = $false; $form.Controls.Add($btnSim)
$btnExec = New-Button 'EXECUTAR (gravar)' 660 512 160 32; $btnExec.Enabled = $false; $btnExec.BackColor = [System.Drawing.Color]::MistyRose; $form.Controls.Add($btnExec)
$btnSave = New-Button 'Salvar log...' 830 512 120 32; $btnSave.Enabled = $false; $form.Controls.Add($btnSave)
$chkItems = New-Check 'Listar cada objeto movido no log' 965 517 250 $false; $form.Controls.Add($chkItems)
$lblStatus = New-Label 'Desconectado.' 490 550 730; $lblStatus.ForeColor = [System.Drawing.Color]::DimGray; $form.Controls.Add($lblStatus)

$dgLog = New-Object System.Windows.Forms.DataGridView; $dgLog.Location = New-Object System.Drawing.Point(490, 572); $dgLog.Size = New-Object System.Drawing.Size(730, 200)
$dgLog.ReadOnly = $true; $dgLog.AllowUserToAddRows = $false; $dgLog.RowHeadersVisible = $false; $dgLog.AutoSizeColumnsMode = 'Fill'; $dgLog.SelectionMode = 'FullRowSelect'
$form.Controls.Add($dgLog)
$script:LogTable = $null

# ----------------------------------------------------------------------------- handlers
function Set-Status { param($t, $err = $false) $lblStatus.Text = $t; $lblStatus.ForeColor = $(if ($err) { [System.Drawing.Color]::Firebrick } else { [System.Drawing.Color]::DimGray }); [System.Windows.Forms.Application]::DoEvents() }
function Show-Error { param($msg) [void][System.Windows.Forms.MessageBox]::Show($msg, 'Move TA', 'OK', 'Error') }

$btnConnect.add_Click({
    try {
        if ($script:Conn) { try { $script:Conn.Close() } catch {} }
        $cs = New-ConnString $tbServer.Text $tbDb.Text $rbWin.Checked $tbUser.Text $tbPass.Text
        $script:Conn = New-Object System.Data.SqlClient.SqlConnection $cs
        $script:Conn.Open()
        $ver = Invoke-SqlScalar "SELECT System_DBVersion + ' / ' + System_Customer FROM ST_System"
        if (-not (Test-Procs)) {
            $r = [System.Windows.Forms.MessageBox]::Show("As procedures usp_IPS_MoveTA / usp_IPS_RetireTA nao existem em $($tbDb.Text).`nInstalar agora a partir de usp_IPS_MoveTA.sql?", 'Move TA', 'YesNo', 'Question')
            if ($r -ne 'Yes') { Set-Status 'Procedures nao instaladas. Instale usp_IPS_MoveTA.sql e conecte de novo.' $true; return }
            Install-Procs
        }
        $cbPlant.Items.Clear()
        $plants = Load-Plants
        foreach ($p in $plants.Rows) { [void]$cbPlant.Items.Add("$($p.Plant_Id)  -  $($p.Plant_Name)") }
        $cbPlant.Enabled = $true
        Set-Status "Conectado: IPS $ver. Escolha a planta."
    } catch { Show-Error $_.Exception.Message; Set-Status 'Falha na conexao.' $true }
})

function Get-PlantId { if ($cbPlant.SelectedItem) { return ([string]$cbPlant.SelectedItem).Split(' ')[0] } return $null }

$cbPlant.add_SelectedIndexChanged({
    try {
        $plant = Get-PlantId; if (-not $plant) { return }
        Set-Status "Carregando TAs da planta $plant..."
        Load-Tree $plant
        $tv.BeginUpdate(); $tv.Nodes.Clear()
        $nodes = @{}
        foreach ($id in ($script:Tree.Keys | Sort-Object { $_.Length }, { $_ })) {
            $n = New-Object System.Windows.Forms.TreeNode; $n.Text = "$id   $($script:Tree[$id].Name)"; $n.Name = $id; $n.Tag = $id
            $nodes[$id] = $n
            $f = $script:Tree[$id].Father
            if ($f -and $nodes.ContainsKey($f)) { [void]$nodes[$f].Nodes.Add($n) } else { [void]$tv.Nodes.Add($n) }
        }
        $tv.EndUpdate()
        $script:Mapping = @(); $dgMap.Rows.Clear(); $lblMode.Text = ''; $btnDest.Enabled = $true; $btnSim.Enabled = $false; $btnExec.Enabled = $false
        Set-Status "Planta $plant (classe $($script:FuncClass)): $($script:Tree.Count) TAs. Marque as TAs de origem e clique em 'Definir destino...'."
    } catch { Show-Error $_.Exception.Message }
})

$btnFilter.add_Click({
    $q = $tbFilter.Text.Trim(); if (-not $q) { return }
    $found = $tv.Nodes.Find($q, $true)
    if ($found.Count -eq 0) { foreach ($k in $script:Tree.Keys) { if ($k -like "$q*" -or $script:Tree[$k].Name -like "*$q*") { $found = $tv.Nodes.Find($k, $true); if ($found.Count -gt 0) { break } } } }
    if ($found.Count -gt 0) { $tv.SelectedNode = $found[0]; $found[0].EnsureVisible(); $tv.Focus() } else { Set-Status "'$q' nao encontrado." $true }
})

function Set-ChildrenChecked { param($node, [bool]$v) foreach ($c in $node.Nodes) { $c.Checked = $v; Set-ChildrenChecked $c $v } }
function Get-CheckedIds { param($nodes) $out = @(); foreach ($n in $nodes) { if ($n.Checked) { $out += [string]$n.Tag }; $out += Get-CheckedIds $n.Nodes }; return ,$out }
$tv.add_AfterCheck({ param($s, $e)
    if ($script:Checking) { return }
    $script:Checking = $true
    try { if ($chkKids.Checked) { Set-ChildrenChecked $e.Node $e.Node.Checked } } finally { $script:Checking = $false }
    $lblSel.Text = "$((Get-CheckedIds $tv.Nodes).Count) selecionadas"
})

function Show-DestDialog {
    <# Pergunta como definir o destino. Devolve @{Mode='prefix'|'custom'; OldPrefix; NewPrefix; TopMap} ou $null #>
    param([string[]]$Top)
    $d = New-Object System.Windows.Forms.Form; $d.Text = 'Definir destino'; $d.Size = New-Object System.Drawing.Size(720, 520); $d.StartPosition = 'CenterParent'; $d.FormBorderStyle = 'FixedDialog'; $d.MaximizeBox = $false
    $d.Font = $form.Font
    $lbl = New-Label "Como definir o destino para as $($Top.Count) TA(s) de topo selecionadas (os subaccounts seguem o mesmo prefixo)?" 15 12 680; $d.Controls.Add($lbl)
    $rb1 = New-Object System.Windows.Forms.RadioButton; $rb1.Text = 'Apenas trocar o inicio do codigo (prefixo)'; $rb1.Location = New-Object System.Drawing.Point(15, 40); $rb1.Size = New-Object System.Drawing.Size(400, 22); $rb1.Checked = $true; $d.Controls.Add($rb1)
    $d.Controls.Add((New-Label 'De:' 40 68 30)); $tbOld = New-TextBox 75 65 100; $d.Controls.Add($tbOld)
    $d.Controls.Add((New-Label 'Para:' 190 68 40)); $tbNew = New-TextBox 235 65 100; $d.Controls.Add($tbNew)
    $lblEx = New-Label '' 350 68 340; $lblEx.ForeColor = [System.Drawing.Color]::DimGray; $d.Controls.Add($lblEx)
    # sugere o prefixo comum ate o ultimo separador
    $common = $Top[0]; foreach ($t in $Top) { $i = 0; while ($i -lt [Math]::Min($common.Length, $t.Length) -and $common[$i] -eq $t[$i]) { $i++ }; $common = $common.Substring(0, $i) }
    $lastSep = $common.LastIndexOf('.'); if ($lastSep -ge 0) { $common = $common.Substring(0, $lastSep + 1) }
    $tbOld.Text = $common
    $rb2 = New-Object System.Windows.Forms.RadioButton; $rb2.Text = 'Escolher o codigo (e opcionalmente o nome) de cada TA de topo'; $rb2.Location = New-Object System.Drawing.Point(15, 100); $rb2.Size = New-Object System.Drawing.Size(500, 22); $d.Controls.Add($rb2)
    $dg = New-Object System.Windows.Forms.DataGridView; $dg.Location = New-Object System.Drawing.Point(15, 128); $dg.Size = New-Object System.Drawing.Size(680, 300); $dg.AllowUserToAddRows = $false; $dg.RowHeadersVisible = $false; $dg.AutoSizeColumnsMode = 'Fill'; $dg.Enabled = $false
    foreach ($c in @(@('Origem',80,$true),@('Nome',200,$true),@('Novo codigo',90,$false),@('Novo nome (opcional)',200,$false))) { $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $col.HeaderText = $c[0]; $col.Name = $c[0]; $col.FillWeight = $c[1]; $col.ReadOnly = $c[2]; [void]$dg.Columns.Add($col) }
    foreach ($t in $Top) { [void]$dg.Rows.Add($t, $script:Tree[$t].Name, '', '') }
    $d.Controls.Add($dg)
    $upd = { $en = $rb1.Checked; $tbOld.Enabled = $en; $tbNew.Enabled = $en; $dg.Enabled = -not $en
             if ($en -and $Top[0].StartsWith($tbOld.Text) -and $tbOld.Text.Length -gt 0) { $lblEx.Text = "ex.: $($Top[0]) -> $($tbNew.Text + $Top[0].Substring($tbOld.Text.Length))" } else { $lblEx.Text = '' } }
    $rb1.add_CheckedChanged($upd); $tbOld.add_TextChanged($upd); $tbNew.add_TextChanged($upd); & $upd
    $ok = New-Button 'OK' 500 440 90; $ok.DialogResult = 'OK'; $d.Controls.Add($ok); $d.AcceptButton = $ok
    $cancel = New-Button 'Cancelar' 600 440 90; $cancel.DialogResult = 'Cancel'; $d.Controls.Add($cancel); $d.CancelButton = $cancel
    if ($d.ShowDialog($form) -ne 'OK') { return $null }
    $topMap = @{}
    if ($rb1.Checked) {
        $old = $tbOld.Text.Trim(); $new = $tbNew.Text.Trim()
        if (-not $old -or -not $new) { Show-Error 'Informe os prefixos "De" e "Para".'; return $null }
        foreach ($t in $Top) { if ($t.StartsWith($old)) { $topMap[$t] = @{ To = $new + $t.Substring($old.Length); ToName = $null } } else { $topMap[$t] = @{ To = $null; ToName = $null } } }
        return @{ Mode = 'prefix'; OldPrefix = $old; NewPrefix = $new; TopMap = $topMap }
    }
    foreach ($r in $dg.Rows) { $to = [string]$r.Cells[2].Value; $nm = [string]$r.Cells[3].Value; $topMap[[string]$r.Cells[0].Value] = @{ To = $(if ($to) { $to.Trim() } else { $null }); ToName = $(if ($nm) { $nm.Trim() } else { $null }) } }
    return @{ Mode = 'custom'; TopMap = $topMap }
}

function Refresh-MapGrid {
    $dgMap.Rows.Clear()
    foreach ($m in $script:Mapping) {
        $i = $dgMap.Rows.Add($m.From, $m.FromName, $m.To, $m.ToName, $m.Status, $m.Links)
        $row = $dgMap.Rows[$i]
        if ($m.Status -like 'INVALIDO*') { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose }
        elseif ($m.Status -like 'MERGE*') { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LemonChiffon }
        if ($m.IsTop) { $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($form.Font, [System.Drawing.FontStyle]::Bold) }
    }
    $bad = @($script:Mapping | Where-Object { $_.Status -like 'INVALIDO*' }).Count
    $merge = @($script:Mapping | Where-Object { $_.Status -like 'MERGE*' }).Count
    $btnSim.Enabled = ($script:Mapping.Count -gt 0 -and $bad -eq 0); $btnExec.Enabled = $false
    Set-Status "$($script:Mapping.Count) pares. $bad invalidos, $merge merges (destino ja existe: as ligacoes sao somadas as da TA existente). Clique em Simular." ($bad -gt 0)
}

$btnDest.add_Click({
    $sel = Get-CheckedIds $tv.Nodes
    if ($sel.Count -eq 0) { Show-Error 'Marque pelo menos uma TA na arvore.'; return }
    if ($chkKids.Checked) { $all = @{}; foreach ($s in $sel) { $all[$s] = 1; foreach ($d in (Get-Descendants $s)) { $all[$d] = 1 } }; $sel = @($all.Keys) }
    $top = Get-TopLevel $sel
    $res = Show-DestDialog $top
    if (-not $res) { return }
    $script:Mapping = Build-Mapping $sel $res.TopMap
    $lblMode.Text = $(if ($res.Mode -eq 'prefix') { "modo: trocar prefixo '$($res.OldPrefix)' -> '$($res.NewPrefix)'" } else { 'modo: codigos escolhidos manualmente' })
    Refresh-MapGrid
})

function Get-Options { return @{ WH = $chkWH.Checked; PO = $chkPO.Checked; Mat = $chkMat.Checked; Other = $chkOther.Checked; Retire = $chkRetire.Checked; RetireClass = $chkRetireClass.Checked; ListItems = $chkItems.Checked } }
function Show-Log { param($log)
    $script:LogTable = $log
    $dgLog.DataSource = $null; $dgLog.AutoGenerateColumns = $true; $dgLog.DataSource = $log
    foreach ($row in $dgLog.Rows) {
        $a = [string]$row.Cells['Acao'].Value
        if ($a -in 'ERRO','ATENCAO','VERIFICAR') { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose }
        elseif ($a -in 'criada','reativada','desativada','COMMIT') { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::Honeydew }
        elseif ($a -in 'MANTIDA','ROLLBACK') { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LemonChiffon }
    }
    $btnSave.Enabled = $true
}

$btnSim.add_Click({
    try {
        Set-Status 'Simulando (transacao com rollback)...'
        $script:LastError = $null
        $log = Run-Moves $script:Mapping (Get-PlantId) $false (Get-Options)
        Show-Log $log
        if ($script:LastError) { Set-Status "Simulacao falhou: $($script:LastError)" $true; $btnExec.Enabled = $false }
        else { Set-Status 'Simulacao OK (nada gravado). Revise o log; se estiver certo, clique em EXECUTAR.'; $btnExec.Enabled = $true }
    } catch { Show-Error $_.Exception.Message }
})
$btnExec.add_Click({
    $plant = Get-PlantId
    $msg = "Vai GRAVAR a realocacao de $($script:Mapping.Count) TA(s) na planta $plant, banco $($tbDb.Text) em $($tbServer.Text).`n`nTudo roda numa unica transacao (ou grava tudo, ou nada).`nFaca backup antes se for producao.`n`nConfirmar?"
    if ([System.Windows.Forms.MessageBox]::Show($msg, 'Move TA - confirmar', 'YesNo', 'Warning') -ne 'Yes') { return }
    try {
        Set-Status 'Executando...'
        $script:LastError = $null
        $log = Run-Moves $script:Mapping $plant $true (Get-Options)
        Show-Log $log
        if ($script:LastError) { Set-Status "ERRO: $($script:LastError)" $true }
        else { Set-Status 'Concluido e gravado. Recarregue a planta para ver a nova arvore.'; $btnExec.Enabled = $false; $btnSim.Enabled = $false }
    } catch { Show-Error $_.Exception.Message }
})
$btnSave.add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog; $sfd.Filter = 'CSV|*.csv'; $sfd.FileName = "MoveTA_$(Get-Date -Format yyyyMMdd_HHmm).csv"
    if ($sfd.ShowDialog() -eq 'OK') { $script:LogTable | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ';'; Set-Status "Log salvo em $($sfd.FileName)" }
})

$form.add_FormClosed({ if ($script:Conn) { try { $script:Conn.Close() } catch {} } })

if ($UiTest) {
    # exercita os handlers da interface sem interacao (uso interno de validacao)
    $form.add_Shown({
        try {
            $btnConnect.PerformClick(); Write-Host "[ui] $($lblStatus.Text)"
            $idx = -1; for ($i = 0; $i -lt $cbPlant.Items.Count; $i++) { if (([string]$cbPlant.Items[$i]).StartsWith($TestPlant)) { $idx = $i } }
            $cbPlant.SelectedIndex = $idx; Write-Host "[ui] $($lblStatus.Text)  nos raiz=$($tv.Nodes.Count)"
            $tbFilter.Text = $TestSources[0]; $btnFilter.PerformClick(); Write-Host "[ui] filtro -> selecionado: $($tv.SelectedNode.Name)"
            foreach ($s in $TestSources) { $n = $tv.Nodes.Find($s, $true); if ($n.Count -gt 0) { $n[0].Checked = $true } }
            Write-Host "[ui] $($lblSel.Text)"
            $sel = Get-CheckedIds $tv.Nodes; $top = Get-TopLevel $sel
            $topMap = @{}; foreach ($t in $top) { $topMap[$t] = @{ To = $TestNewPrefix + $t.Substring($TestOldPrefix.Length); ToName = $null } }
            $script:Mapping = Build-Mapping $sel $topMap; $lblMode.Text = 'modo: teste'; Refresh-MapGrid
            Write-Host "[ui] grid mapeamento: $($dgMap.Rows.Count) linhas; status: $($lblStatus.Text)"
            $btnSim.PerformClick(); Write-Host "[ui] simular -> $($lblStatus.Text); log=$($dgLog.Rows.Count) linhas; EXECUTAR habilitado=$($btnExec.Enabled)"
            if ($ScreenshotDir) {
                foreach ($s in $TestSources) { $n = $tv.Nodes.Find($s, $true); if ($n.Count -gt 0) { $n[0].Expand(); $n[0].EnsureVisible() } }
                [System.Windows.Forms.Application]::DoEvents()
                $bmp = New-Object System.Drawing.Bitmap($form.Width, $form.Height)
                $form.DrawToBitmap($bmp, (New-Object System.Drawing.Rectangle(0, 0, $form.Width, $form.Height)))
                $p1 = Join-Path $ScreenshotDir 'MoveTA_tela_principal.png'; $bmp.Save($p1, [System.Drawing.Imaging.ImageFormat]::Png); Write-Host "[ui] print: $p1"
                # dialogo de destino: abre, captura e cancela
                $tmr = New-Object System.Windows.Forms.Timer; $tmr.Interval = 700
                $tmr.add_Tick({ $tmr.Stop(); $dlg = [System.Windows.Forms.Form]::ActiveForm
                    if ($dlg -and $dlg -ne $form) { $b2 = New-Object System.Drawing.Bitmap($dlg.Width, $dlg.Height); $dlg.DrawToBitmap($b2, (New-Object System.Drawing.Rectangle(0, 0, $dlg.Width, $dlg.Height)))
                        $p2 = Join-Path $ScreenshotDir 'MoveTA_dialogo_destino.png'; $b2.Save($p2, [System.Drawing.Imaging.ImageFormat]::Png); Write-Host "[ui] print: $p2"; $dlg.DialogResult = 'Cancel' } })
                $tmr.Start(); [void](Show-DestDialog (Get-TopLevel (Get-CheckedIds $tv.Nodes)))
            }
        } catch { Write-Host "[ui] ERRO: $($_.Exception.Message)`n$($_.ScriptStackTrace)" }
        $form.Close()
    })
}
[void]$form.ShowDialog()
