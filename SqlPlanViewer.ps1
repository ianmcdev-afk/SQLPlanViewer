<#
.SYNOPSIS
    SQL Plan Viewer - friendly GUI for SQL Server execution plans.

.DESCRIPTION
    * Open .sqlplan / showplan XML files (or drag & drop them onto the window)
    * Paste plan XML from the clipboard
    * Get an estimated or actual plan straight from SQL Server
    * Graphical plan (zoom, export PNG), operator tree, top operators,
      warnings & findings, missing indexes, parameters / statement summary
    * Per-operator property panel

    The console window is hidden automatically; pass -ShowConsole to keep it
    (handy if the app fails to start).

    Works in Windows PowerShell 5.1 and PowerShell 7 on Windows.
    Passwords are never saved. Actual plans EXECUTE the query (optionally
    wrapped in a transaction that is rolled back).

.PARAMETER PlanFile
    Optional .sqlplan file to open at start-up.

.EXAMPLE
    .\SqlPlanViewer.ps1
    .\SqlPlanViewer.ps1 C:\temp\slow-query.sqlplan
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$PlanFile,

    # Keep the PowerShell console window visible (useful for troubleshooting).
    [switch]$ShowConsole,

    # Internal: set on the hidden/STA copy the script starts for itself.
    [Parameter(DontShow)]
    [switch]$Relaunched
)

$ErrorActionPreference = 'Stop'

# WinForms needs an STA thread, and the console window is hidden by default:
# start a hidden STA copy of this script and let this one exit.
# (Skipped when compiled to an .exe - there is no script path then, and the exe is built STA/no-console.)
$isSta = ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq [System.Threading.ApartmentState]::STA)
if ($PSCommandPath -and -not $Relaunched -and (-not $isSta -or -not $ShowConsole)) {
    $relaunch = @('-STA', '-NoProfile')
    if (-not $ShowConsole) { $relaunch += @('-WindowStyle', 'Hidden') }
    $relaunch += @('-File', "`"$PSCommandPath`"", '-Relaunched')
    if ($ShowConsole) { $relaunch += '-ShowConsole' }
    if ($PlanFile) { $relaunch += @('-PlanFile', "`"$PlanFile`"") }
    $sp = @{ FilePath = (Get-Process -Id $PID).Path; ArgumentList = $relaunch }
    if (-not $ShowConsole) { $sp.WindowStyle = 'Hidden' }
    Start-Process @sp
    return
}

#region Plan model (no GUI dependencies)

$script:NsUri = 'http://schemas.microsoft.com/sqlserver/2004/07/showplan'
$script:Ns = New-Object System.Xml.XmlNamespaceManager((New-Object System.Xml.NameTable))
$script:Ns.AddNamespace('sp', $script:NsUri)

# Layout constants
$script:NodeW  = 190
$script:NodeH  = 76
$script:GapX   = 70
$script:GapY   = 18
$script:Margin = 24
$script:LayoutY = 0.0

function Get-Attr {
    param($El, [string]$Name)
    if ($null -eq $El) { return '' }
    return $El.GetAttribute($Name)
}

function Get-Num {
    param($Value)
    $d = 0.0
    $s = [string]$Value
    if ($s -and [double]::TryParse($s, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }
    return 0.0
}

function Remove-Brackets {
    param([string]$Text)
    return ($Text -replace '[\[\]]', '')
}

function Format-Count {
    param([double]$V)
    if ($V -ge 1e9) { return ('{0:0.##}B' -f ($V / 1e9)) }
    if ($V -ge 1e6) { return ('{0:0.##}M' -f ($V / 1e6)) }
    if ($V -ge 1e4) { return ('{0:0.#}K' -f ($V / 1e3)) }
    if ($V -eq [math]::Floor($V)) { return ('{0:N0}' -f $V) }
    return ('{0:0.##}' -f $V)
}

function Get-WarningTexts {
    param($W)
    $out = New-Object System.Collections.Generic.List[string]
    if ($null -eq $W) { return }
    foreach ($a in $W.Attributes) {
        if ($a.Value -eq 'true' -or $a.Value -eq '1') {
            $out.Add(($a.Name -creplace '([a-z])([A-Z])', '$1 $2'))
        }
    }
    foreach ($c in $W.ChildNodes) {
        if ($c -isnot [System.Xml.XmlElement]) { continue }
        switch ($c.LocalName) {
            'ColumnsWithNoStatistics' {
                $cols = @($c.SelectNodes('.//sp:ColumnReference', $script:Ns) | ForEach-Object { Remove-Brackets ($_.GetAttribute('Column')) })
                $out.Add('Columns with no statistics: ' + ($cols -join ', '))
            }
            'PlanAffectingConvert' {
                $out.Add('Implicit conversion affects ' + $c.GetAttribute('ConvertIssue') + ': ' + $c.GetAttribute('Expression'))
            }
            default {
                $attrs = @($c.Attributes | ForEach-Object { '{0}={1}' -f $_.Name, $_.Value }) -join ', '
                $txt = ($c.LocalName -creplace '([a-z])([A-Z])', '$1 $2')
                if ($attrs) { $txt += " ($attrs)" }
                $out.Add($txt)
            }
        }
    }
    return $out
}

function ConvertTo-PlanNode {
    param([System.Xml.XmlElement]$Rel, $Parent, [int]$Depth = 0)
    if ($Depth -gt 200) { throw 'Plan is nested more than 200 operators deep - refusing to load it (malformed or hostile file).' }

    $phys = Get-Attr $Rel 'PhysicalOp'
    $logi = Get-Attr $Rel 'LogicalOp'
    $obj  = $Rel.SelectSingleNode('./*/sp:Object', $script:Ns)
    $kind = ''; $short = ''; $full = ''
    if ($obj) {
        $tbl = Get-Attr $obj 'Table'
        $idx = Get-Attr $obj 'Index'
        $kind = Get-Attr $obj 'IndexKind'
        $shortRaw = $tbl
        if ($idx) { $shortRaw = "$tbl.$idx" }
        $short = Remove-Brackets $shortRaw
        $full = (@((Get-Attr $obj 'Database'), (Get-Attr $obj 'Schema'), $tbl, $idx) | Where-Object { $_ }) -join '.'
    }

    $isLookup = $false
    $scan = $Rel.SelectSingleNode('./sp:IndexScan', $script:Ns)
    if ($scan -and ((Get-Attr $scan 'Lookup') -in @('1', 'true'))) { $isLookup = $true }
    if ($phys -match 'Key Lookup|RID Lookup') { $isLookup = $true }

    $disp = $phys
    if ($kind -and $phys -match '^Index (Seek|Scan)$') { $disp = "$phys ($kind)" }
    if ($isLookup) {
        if ($phys -match 'RID' -or $kind -eq 'Heap') { $disp = 'RID Lookup (Heap)' }
        elseif ($kind) { $disp = "Key Lookup ($kind)" }
        else { $disp = 'Key Lookup' }
    }
    if (-not $disp) { $disp = $logi }

    $n = [pscustomobject]@{
        Id          = [int](Get-Num (Get-Attr $Rel 'NodeId'))
        PhysicalOp  = $phys
        LogicalOp   = $logi
        Display     = $disp
        ObjectShort = $short
        ObjectFull  = $full
        IsLookup    = $isLookup
        EstRows     = (Get-Num (Get-Attr $Rel 'EstimateRows'))
        EstRowsRead = (Get-Num (Get-Attr $Rel 'EstimatedRowsRead'))
        SubtreeCost = (Get-Num (Get-Attr $Rel 'EstimatedTotalSubtreeCost'))
        SelfCost    = 0.0
        CostPct     = 0.0
        HasRuntime  = $false
        ActualRows  = 0.0
        ActualExec  = 0.0
        ActualElapsed = 0.0
        ActualCpu   = 0.0
        ActualReads = 0.0
        Warnings    = (New-Object System.Collections.Generic.List[string])
        Children    = (New-Object System.Collections.Generic.List[object])
        Parent      = $Parent
        X           = 0.0
        Y           = 0.0
        TreeNode    = $null
        Xml         = $Rel
    }

    $rt = $Rel.SelectNodes('./sp:RunTimeInformation/sp:RunTimeCountersPerThread', $script:Ns)
    if ($rt.Count -gt 0) {
        $n.HasRuntime = $true
        foreach ($t in $rt) {
            $n.ActualRows  += (Get-Num (Get-Attr $t 'ActualRows'))
            $n.ActualExec  += (Get-Num (Get-Attr $t 'ActualExecutions'))
            $n.ActualCpu   += (Get-Num (Get-Attr $t 'ActualCPUms'))
            $n.ActualReads += (Get-Num (Get-Attr $t 'ActualLogicalReads'))
            $n.ActualElapsed = [math]::Max($n.ActualElapsed, (Get-Num (Get-Attr $t 'ActualElapsedms')))
        }
    }

    foreach ($w in @(Get-WarningTexts ($Rel.SelectSingleNode('./sp:Warnings', $script:Ns)))) { $n.Warnings.Add($w) }

    $childSum = 0.0
    foreach ($c in $Rel.SelectNodes('./*/sp:RelOp', $script:Ns)) {
        $child = ConvertTo-PlanNode -Rel $c -Parent $n -Depth ($Depth + 1)
        $n.Children.Add($child)
        $childSum += $child.SubtreeCost
    }
    $n.SelfCost = [math]::Max(0.0, $n.SubtreeCost - $childSum)
    return $n
}

function Get-PlanNodes {
    param($Root)
    $list = New-Object System.Collections.Generic.List[object]
    $stack = New-Object System.Collections.Generic.Stack[object]
    $stack.Push($Root)
    while ($stack.Count -gt 0) {
        $cur = $stack.Pop()
        $list.Add($cur)
        for ($i = $cur.Children.Count - 1; $i -ge 0; $i--) { $stack.Push($cur.Children[$i]) }
    }
    return , $list
}

function Read-PlanXml {
    param([string]$Text)

    $settings = New-Object System.Xml.XmlReaderSettings
    $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver = $null
    $settings.MaxCharactersInDocument = 100000000
    $doc = New-Object System.Xml.XmlDocument
    $sr = New-Object System.IO.StringReader($Text.Trim())
    $xr = [System.Xml.XmlReader]::Create($sr, $settings)
    try { $doc.Load($xr) } finally { $xr.Close(); $sr.Dispose() }

    if ($doc.DocumentElement.LocalName -ne 'ShowPlanXML') {
        throw 'This does not look like SQL Server showplan XML (root element ShowPlanXML not found).'
    }

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($s in $doc.SelectNodes('//sp:StmtSimple[sp:QueryPlan/sp:RelOp]', $script:Ns)) {
        $qp = $s.SelectSingleNode('./sp:QueryPlan', $script:Ns)
        $rootNode = ConvertTo-PlanNode -Rel $qp.SelectSingleNode('./sp:RelOp', $script:Ns) -Parent $null
        $nodes = Get-PlanNodes $rootNode

        $den = 1.0
        if ($rootNode.SubtreeCost -gt 0) { $den = $rootNode.SubtreeCost }
        foreach ($n in $nodes) { $n.CostPct = $n.SelfCost / $den * 100.0 }

        $cost = Get-Num (Get-Attr $s 'StatementSubTreeCost')
        if ($cost -le 0) { $cost = $rootNode.SubtreeCost }

        $txtFull = (Get-Attr $s 'StatementText').Trim()
        $oneLine = ($txtFull -replace '\s+', ' ').Trim()

        # Missing indexes
        $mis = New-Object System.Collections.Generic.List[object]
        foreach ($g in $qp.SelectNodes('./sp:MissingIndexes/sp:MissingIndexGroup', $script:Ns)) {
            $impact = Get-Num (Get-Attr $g 'Impact')
            foreach ($mi in $g.SelectNodes('./sp:MissingIndex', $script:Ns)) {
                $eq = @(); $ineq = @(); $inc = @()
                foreach ($cg in $mi.SelectNodes('./sp:ColumnGroup', $script:Ns)) {
                    $cols = @($cg.SelectNodes('./sp:Column', $script:Ns) | ForEach-Object { $_.GetAttribute('Name') })
                    switch (Get-Attr $cg 'Usage') {
                        'EQUALITY'   { $eq += $cols }
                        'INEQUALITY' { $ineq += $cols }
                        'INCLUDE'    { $inc += $cols }
                    }
                }
                $tableName = (@((Get-Attr $mi 'Database'), (Get-Attr $mi 'Schema'), (Get-Attr $mi 'Table')) | Where-Object { $_ }) -join '.'
                $key = @($eq + $ineq)
                $ixName = 'IX_' + (Remove-Brackets (Get-Attr $mi 'Table')) + '_' + ((Remove-Brackets ($key -join '_')) -replace '[^A-Za-z0-9_]', '')
                if ($ixName.Length -gt 100) { $ixName = $ixName.Substring(0, 100) }
                $script = ''
                if ($key.Count -gt 0) {
                    $script = "CREATE NONCLUSTERED INDEX [$ixName]`r`n    ON $tableName ($($key -join ', '))"
                    if ($inc.Count -gt 0) { $script += "`r`n    INCLUDE ($($inc -join ', '))" }
                    $script += ';'
                }
                $mis.Add([pscustomobject]@{
                    Impact  = $impact
                    Table   = $tableName
                    Key     = ($key -join ', ')
                    Include = ($inc -join ', ')
                    Script  = $script
                })
            }
        }

        $params = @()
        foreach ($p in $qp.SelectNodes('./sp:ParameterList/sp:ColumnReference', $script:Ns)) {
            $params += [pscustomobject]@{
                Name     = (Get-Attr $p 'Column')
                Type     = (Get-Attr $p 'ParameterDataType')
                Compiled = (Get-Attr $p 'ParameterCompiledValue')
                Runtime  = (Get-Attr $p 'ParameterRuntimeValue')
            }
        }

        $planWarn = @(Get-WarningTexts ($qp.SelectSingleNode('./sp:Warnings', $script:Ns)))

        $result.Add([pscustomobject]@{
            Index          = $result.Count + 1
            Type           = (Get-Attr $s 'StatementType')
            Text           = $txtFull
            OneLine        = $oneLine
            Cost           = $cost
            CostPct        = 0.0
            EstRows        = (Get-Num (Get-Attr $s 'StatementEstRows'))
            Root           = $rootNode
            Nodes          = $nodes
            Warnings       = $planWarn
            MissingIndexes = $mis
            Params         = $params
            StmtXml        = $s
            QpXml          = $qp
            SourceIndex    = 0
            MaxX           = 0.0
            MaxY           = 0.0
        })
    }
    return $result
}

function Set-PlanLayout {
    param($N, [int]$Depth)
    $N.X = $script:Margin + $Depth * ($script:NodeW + $script:GapX)
    if ($N.Children.Count -eq 0) {
        $N.Y = $script:LayoutY
        $script:LayoutY += $script:NodeH + $script:GapY
    } else {
        foreach ($c in $N.Children) { Set-PlanLayout $c ($Depth + 1) }
        $N.Y = ($N.Children[0].Y + $N.Children[$N.Children.Count - 1].Y) / 2
    }
}

function Get-NodeDetails {
    param($N)
    $d = New-Object System.Collections.Generic.List[object]
    $add = {
        param($g, $k, $v)
        if ($null -ne $v -and "$v" -ne '') { $d.Add([pscustomobject]@{ G = $g; N = $k; V = "$v" }) }
    }
    $x = $N.Xml

    $par = ''
    if ((Get-Attr $x 'Parallel') -in @('1', 'true')) { $par = 'Yes' }
    $actMode = ''
    $firstRt = $x.SelectSingleNode('./sp:RunTimeInformation/sp:RunTimeCountersPerThread', $script:Ns)
    if ($firstRt) { $actMode = Get-Attr $firstRt 'ActualExecutionMode' }

    & $add 'Operator' 'Physical operator' $N.PhysicalOp
    & $add 'Operator' 'Logical operator' $N.LogicalOp
    & $add 'Operator' 'Node ID' $N.Id
    & $add 'Operator' 'Object' $N.ObjectFull
    & $add 'Operator' 'Parallel' $par
    & $add 'Operator' 'Estimated execution mode' (Get-Attr $x 'EstimatedExecutionMode')
    & $add 'Operator' 'Actual execution mode' $actMode

    & $add 'Cost' 'Cost (% of statement)' ('{0:0.0}%' -f $N.CostPct)
    & $add 'Cost' 'Operator cost' ('{0:0.000000}' -f $N.SelfCost)
    & $add 'Cost' 'Subtree cost' ('{0:0.000000}' -f $N.SubtreeCost)
    & $add 'Cost' 'Estimated I/O cost' ('{0:0.000000}' -f (Get-Num (Get-Attr $x 'EstimateIO')))
    & $add 'Cost' 'Estimated CPU cost' ('{0:0.000000}' -f (Get-Num (Get-Attr $x 'EstimateCPU')))

    & $add 'Rows' 'Estimated rows (per execution)' (Format-Count $N.EstRows)
    if ($N.EstRowsRead -gt 0) { & $add 'Rows' 'Estimated rows to be read' (Format-Count $N.EstRowsRead) }
    if ($N.HasRuntime) {
        & $add 'Rows' 'Actual rows' ('{0:N0}' -f $N.ActualRows)
        & $add 'Rows' 'Actual executions' ('{0:N0}' -f $N.ActualExec)
        $estTot = $N.EstRows * [math]::Max($N.ActualExec, 1)
        if ($estTot -gt 0) { & $add 'Rows' 'Actual / estimated' ('{0:0.0#}x' -f ($N.ActualRows / $estTot)) }
        $rr = 0.0
        foreach ($t in $x.SelectNodes('./sp:RunTimeInformation/sp:RunTimeCountersPerThread', $script:Ns)) { $rr += (Get-Num (Get-Attr $t 'ActualRowsRead')) }
        if ($rr -gt 0) { & $add 'Rows' 'Actual rows read' ('{0:N0}' -f $rr) }
    }
    & $add 'Rows' 'Average row size' ((Get-Attr $x 'AvgRowSize') + ' B')
    $tc = Get-Attr $x 'TableCardinality'
    if ($tc) { & $add 'Rows' 'Table cardinality' (Format-Count (Get-Num $tc)) }

    if ($N.HasRuntime) {
        & $add 'Runtime' 'Elapsed time' ('{0:N0} ms' -f $N.ActualElapsed)
        & $add 'Runtime' 'CPU time' ('{0:N0} ms' -f $N.ActualCpu)
        if ($N.ActualReads -gt 0) { & $add 'Runtime' 'Logical reads' ('{0:N0}' -f $N.ActualReads) }
    }

    $seek = New-Object System.Collections.Generic.List[string]
    foreach ($e in $x.SelectNodes('./*/sp:SeekPredicates//*[self::sp:Prefix or self::sp:StartRange or self::sp:EndRange]', $script:Ns)) {
        $cols = @($e.SelectNodes('./sp:RangeColumns/sp:ColumnReference', $script:Ns) | ForEach-Object { Remove-Brackets ($_.GetAttribute('Column')) })
        $exprs = @($e.SelectNodes('./sp:RangeExpressions/sp:ScalarOperator', $script:Ns) | ForEach-Object { $_.GetAttribute('ScalarString') })
        $seek.Add(('{0}: {1} {2} {3}' -f $e.LocalName, ($cols -join ', '), (Get-Attr $e 'ScanType'), ($exprs -join ', ')))
    }
    $pred = @($x.SelectNodes('./*/sp:Predicate/sp:ScalarOperator', $script:Ns) | ForEach-Object { $_.GetAttribute('ScalarString') } | Where-Object { $_ })
    $resid = @($x.SelectNodes('./*/sp:ProbeResidual/sp:ScalarOperator', $script:Ns) | ForEach-Object { $_.GetAttribute('ScalarString') } | Where-Object { $_ })
    & $add 'Predicates' 'Seek predicates' ($seek -join ' | ')
    & $add 'Predicates' 'Predicate' ($pred -join ' AND ')
    & $add 'Predicates' 'Probe residual' ($resid -join ' AND ')

    $outCols = @($x.SelectNodes('./sp:OutputList/sp:ColumnReference', $script:Ns) | ForEach-Object {
        $tb = Remove-Brackets ($_.GetAttribute('Alias'))
        if (-not $tb) { $tb = Remove-Brackets ($_.GetAttribute('Table')) }
        $cn = Remove-Brackets ($_.GetAttribute('Column'))
        if ($tb) { "$tb.$cn" } else { $cn }
    })
    if ($outCols.Count -gt 0) {
        $txt = (($outCols | Select-Object -First 30) -join ', ')
        if ($outCols.Count -gt 30) { $txt += ' ... (+{0} more)' -f ($outCols.Count - 30) }
        & $add 'Predicates' ('Output columns ({0})' -f $outCols.Count) $txt
    }

    foreach ($w in $N.Warnings) { & $add 'Warnings' 'Warning' $w }
    return , $d
}

function Get-Findings {
    param($St)
    $f = New-Object System.Collections.Generic.List[object]
    $add = {
        param($sev, $n, $msg)
        $op = '(statement)'
        if ($n) { $op = '{0} (node {1})' -f $n.Display, $n.Id }
        $f.Add([pscustomobject]@{ Severity = $sev; Operator = $op; Message = $msg; Node = $n })
    }
    foreach ($w in $St.Warnings) { & $add 'Warning' $null $w }
    foreach ($mi in $St.MissingIndexes) {
        & $add 'Warning' $null ('Missing index suggested (estimated impact {0:0.#}%) on {1}' -f $mi.Impact, $mi.Table)
    }
    foreach ($n in $St.Nodes) { foreach ($w in $n.Warnings) { & $add 'Warning' $n $w } }

    foreach ($n in $St.Nodes) {
        if ($n.CostPct -ge 25) {
            & $add 'Info' $n ('Accounts for {0:0}% of the statement cost.' -f $n.CostPct)
        }
        if (-not $n.IsLookup -and $n.PhysicalOp -match '^(Table Scan|Clustered Index Scan|Index Scan)$') {
            $r = [math]::Max($n.EstRows, $n.EstRowsRead)
            if ($r -ge 10000) {
                & $add 'Info' $n ('Scans about {0} rows. Check for a missing index or a non-sargable predicate.' -f (Format-Count $r))
            }
        }
        if ($n.IsLookup) {
            $outer = 0.0
            if ($n.HasRuntime) { $outer = $n.ActualExec }
            elseif ($n.Parent -and $n.Parent.Children.Count -ge 2) { $outer = $n.Parent.Children[0].EstRows }
            if ($outer -ge 1000) {
                & $add 'Info' $n ('Lookup runs about {0} times. A covering index (INCLUDE the looked-up columns) may help.' -f (Format-Count $outer))
            }
        }
        if ($n.HasRuntime -and $n.EstRows -gt 0) {
            $estTot = $n.EstRows * [math]::Max($n.ActualExec, 1)
            $hi = [math]::Max($estTot, $n.ActualRows)
            $ratio = ($n.ActualRows + 1) / ($estTot + 1)
            if ($hi -ge 1000 -and ($ratio -ge 10 -or $ratio -le 0.1)) {
                & $add 'Info' $n ('Row estimate is off: estimated {0}, actual {1}. Check statistics.' -f (Format-Count $estTot), (Format-Count $n.ActualRows))
            }
        }
    }
    return , $f
}

function Get-StatementSummary {
    param($St)
    $sb = New-Object System.Text.StringBuilder
    $line = {
        param($k, $v)
        if ($null -ne $v -and "$v" -ne '') { [void]$sb.AppendLine(('{0,-30}{1}' -f ($k + ':'), $v)) }
    }
    $s = $St.StmtXml
    $q = $St.QpXml
    & $line 'Statement type' $St.Type
    & $line 'Estimated subtree cost' ('{0:0.000000}' -f $St.Cost)
    & $line 'Estimated rows' (Format-Count $St.EstRows)
    & $line 'Optimization level' (Get-Attr $s 'StatementOptmLevel')
    & $line 'Early abort reason' (Get-Attr $s 'StatementOptmEarlyAbortReason')
    & $line 'Cardinality estimator' (Get-Attr $s 'CardinalityEstimationModelVersion')
    & $line 'Query hash' (Get-Attr $s 'QueryHash')
    & $line 'Query plan hash' (Get-Attr $s 'QueryPlanHash')
    & $line 'Degree of parallelism' (Get-Attr $q 'DegreeOfParallelism')
    & $line 'Non-parallel reason' (Get-Attr $q 'NonParallelPlanReason')
    & $line 'Compile time (ms)' (Get-Attr $q 'CompileTime')
    & $line 'Compile CPU (ms)' (Get-Attr $q 'CompileCPU')
    & $line 'Compile memory (KB)' (Get-Attr $q 'CompileMemory')
    & $line 'Cached plan size (KB)' (Get-Attr $q 'CachedPlanSize')
    $mg = $q.SelectSingleNode('./sp:MemoryGrantInfo', $script:Ns)
    if ($mg) {
        & $line 'Memory grant requested (KB)' (Get-Attr $mg 'RequestedMemory')
        & $line 'Memory grant granted (KB)' (Get-Attr $mg 'GrantedMemory')
        & $line 'Memory grant used (KB)' (Get-Attr $mg 'MaxUsedMemory')
    }
    $ts = $q.SelectSingleNode('./sp:QueryTimeStats', $script:Ns)
    if ($ts) {
        & $line 'Elapsed time (ms)' (Get-Attr $ts 'ElapsedTime')
        & $line 'CPU time (ms)' (Get-Attr $ts 'CpuTime')
    }
    foreach ($w in $q.SelectNodes('./sp:WaitStats/sp:Wait', $script:Ns)) {
        & $line ('Wait ' + (Get-Attr $w 'WaitType')) ('{0} ms ({1} waits)' -f (Get-Attr $w 'WaitTimeMs'), (Get-Attr $w 'WaitCount'))
    }
    if (@($St.Params).Count -gt 0) {
        [void]$sb.AppendLine()
        [void]$sb.AppendLine('Parameters')
        foreach ($p in $St.Params) {
            [void]$sb.AppendLine(('  {0} ({1})   compiled = {2}   runtime = {3}' -f $p.Name, $p.Type, $p.Compiled, $p.Runtime))
        }
    }
    [void]$sb.AppendLine()
    [void]$sb.AppendLine('Statement text')
    [void]$sb.AppendLine('--------------')
    [void]$sb.AppendLine($St.Text)
    return $sb.ToString()
}

#endregion

#region GUI

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data
try { Add-Type -AssemblyName System.Data.SqlClient -ErrorAction Stop } catch { }
try { [System.Windows.Forms.Application]::SetUnhandledExceptionMode([System.Windows.Forms.UnhandledExceptionMode]::CatchException) } catch { }
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:Statements  = @()
$script:PlanTexts   = @()
$script:CurrentStmt = $null
$script:SelNode     = $null
$script:HoverNode   = $null
$script:Syncing     = $false
$script:Zoom        = 1.0
$script:LastQuery   = ''
$script:SettingsPath = Join-Path $env:APPDATA 'SqlPlanViewer\settings.json'

$script:FontBold = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$script:FontNorm = New-Object System.Drawing.Font('Segoe UI', 8.5)
$script:Sf = New-Object System.Drawing.StringFormat
$script:Sf.Trimming = [System.Drawing.StringTrimming]::EllipsisCharacter
$script:Sf.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap
$script:SfCenter = New-Object System.Drawing.StringFormat
$script:SfCenter.Alignment = [System.Drawing.StringAlignment]::Center
$script:SfCenter.LineAlignment = [System.Drawing.StringAlignment]::Center

$UI = @{}

# ---------- helpers ----------

function Show-Error {
    param($ErrorRecord)
    $msg = $ErrorRecord.Exception.Message
    [void][System.Windows.Forms.MessageBox]::Show($UI.Form, $msg, 'SQL Plan Viewer', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}

function Show-Info {
    param([string]$Text)
    [void][System.Windows.Forms.MessageBox]::Show($UI.Form, $Text, 'SQL Plan Viewer', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

function New-Control {
    param([string]$Type, [int]$X, [int]$Y, [int]$W, [int]$H, [string]$Text = '')
    $c = New-Object ("System.Windows.Forms.$Type")
    $c.Location = New-Object System.Drawing.Point($X, $Y)
    $c.Size = New-Object System.Drawing.Size($W, $H)
    if ($Text) { $c.Text = $Text }
    return $c
}

function Set-DoubleBuffered {
    param($Control)
    $p = [System.Windows.Forms.Control].GetProperty('DoubleBuffered', [System.Reflection.BindingFlags]'Instance,NonPublic')
    if ($p) { $p.SetValue($Control, $true, $null) }
}

function New-RoundRectPath {
    param([single]$X, [single]$Y, [single]$W, [single]$H, [single]$R)
    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = [single]($R * 2)
    $p.AddArc($X, $Y, $d, $d, [single]180, [single]90)
    $p.AddArc([single]($X + $W - $d), $Y, $d, $d, [single]270, [single]90)
    $p.AddArc([single]($X + $W - $d), [single]($Y + $H - $d), $d, $d, [single]0, [single]90)
    $p.AddArc($X, [single]($Y + $H - $d), $d, $d, [single]90, [single]90)
    $p.CloseFigure()
    return $p
}

function Wait-Task {
    param($Task)
    while (-not $Task.IsCompleted) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 25
    }
    if ($Task.IsFaulted) {
        $ex = $Task.Exception
        if ($ex.InnerException) { throw $ex.InnerException }
        throw $ex
    }
    if ($Task.IsCanceled) { throw 'The operation was cancelled.' }
}

function Get-ReaderXml {
    param($Reader)
    try { return $Reader.GetSqlXml(0).Value } catch { return $Reader.GetString(0) }
}

# ---------- settings (never stores passwords or query text) ----------

function Get-Settings {
    $d = @{ Server = ''; Database = ''; Windows = $true; User = ''; Trust = $false; Actual = $false; Rollback = $true; Timeout = 120 }
    try {
        if (Test-Path -LiteralPath $script:SettingsPath) {
            $j = Get-Content -LiteralPath $script:SettingsPath -Raw | ConvertFrom-Json
            foreach ($k in @($d.Keys)) { if ($null -ne $j.$k) { $d[$k] = $j.$k } }
        }
    } catch { }
    return $d
}

function Save-Settings {
    param($Cfg)
    try {
        $dir = Split-Path -Parent $script:SettingsPath
        if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir) }
        @{
            Server = $Cfg.Server; Database = $Cfg.Database; Windows = $Cfg.Windows; User = $Cfg.User
            Trust = $Cfg.Trust; Actual = $Cfg.Actual; Rollback = $Cfg.Rollback; Timeout = $Cfg.Timeout
        } | ConvertTo-Json | Set-Content -LiteralPath $script:SettingsPath -Encoding UTF8
    } catch { }
}

# ---------- SQL Server access ----------

function Show-QueryDialog {
    param($Preset)
    $s = Get-Settings
    if ($Preset) {
        foreach ($k in @($Preset.Keys)) { if ($k -ne 'Query' -and $k -ne 'Password') { $s[$k] = $Preset[$k] } }
    }
    $f = New-Object System.Windows.Forms.Form
    $f.Text = 'Get execution plan from SQL Server'
    $f.ClientSize = New-Object System.Drawing.Size(780, 650)
    $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'FixedDialog'
    $f.MaximizeBox = $false
    $f.MinimizeBox = $false
    $f.ShowInTaskbar = $false
    $f.Font = $UI.Font

    $f.Controls.Add((New-Control 'Label' 15 18 75 20 'Server'))
    $tbServer = New-Control 'TextBox' 95 15 300 23
    $tbServer.Text = $s.Server
    $f.Controls.Add($tbServer)

    $f.Controls.Add((New-Control 'Label' 410 18 60 20 'Auth'))
    $cbAuth = New-Control 'ComboBox' 470 15 295 23
    $cbAuth.DropDownStyle = 'DropDownList'
    [void]$cbAuth.Items.Add('Windows authentication')
    [void]$cbAuth.Items.Add('SQL Server authentication')
    if ($s.Windows) { $cbAuth.SelectedIndex = 0 } else { $cbAuth.SelectedIndex = 1 }
    $f.Controls.Add($cbAuth)

    $f.Controls.Add((New-Control 'Label' 15 51 75 20 'Database'))
    $tbDb = New-Control 'TextBox' 95 48 300 23
    $tbDb.Text = $s.Database
    $f.Controls.Add($tbDb)

    $f.Controls.Add((New-Control 'Label' 410 51 60 20 'Login'))
    $tbUser = New-Control 'TextBox' 470 48 295 23
    $tbUser.Text = $s.User
    $f.Controls.Add($tbUser)

    $f.Controls.Add((New-Control 'Label' 410 84 60 20 'Password'))
    $tbPass = New-Control 'TextBox' 470 81 295 23
    $tbPass.UseSystemPasswordChar = $true
    if ($Preset -and $Preset.Password) { $tbPass.Text = [string]$Preset.Password }
    $f.Controls.Add($tbPass)

    $chkTrust = New-Control 'CheckBox' 95 82 300 22 'Trust server certificate (no validation)'
    $chkTrust.Checked = [bool]$s.Trust
    $f.Controls.Add($chkTrust)

    $f.Controls.Add((New-Control 'Label' 15 117 75 20 'Timeout (s)'))
    $nud = New-Control 'NumericUpDown' 95 114 80 23
    $nud.Minimum = 5
    $nud.Maximum = 36000
    $nud.Value = [decimal][math]::Max(5, [int]$s.Timeout)
    $f.Controls.Add($nud)

    $gb = New-Control 'GroupBox' 15 145 750 108 'Plan type'
    $rbEst = New-Control 'RadioButton' 15 24 700 22 'Estimated plan  -  the query is NOT executed (safe)'
    $rbAct = New-Control 'RadioButton' 15 48 700 22 'Actual plan  -  the query IS executed (results are discarded)'
    $chkRb = New-Control 'CheckBox' 38 74 680 22 'Wrap in a transaction and roll back afterwards (recommended for INSERT/UPDATE/DELETE)'
    $rbEst.Checked = -not [bool]$s.Actual
    $rbAct.Checked = [bool]$s.Actual
    $chkRb.Checked = [bool]$s.Rollback
    $chkRb.Enabled = $rbAct.Checked
    $gb.Controls.Add($rbEst); $gb.Controls.Add($rbAct); $gb.Controls.Add($chkRb)
    $f.Controls.Add($gb)

    $f.Controls.Add((New-Control 'Label' 15 262 600 20 'Query  (use GO on its own line to separate batches; pick the database above rather than using USE)'))
    $tbQuery = New-Control 'TextBox' 15 285 750 310
    $tbQuery.Multiline = $true
    $tbQuery.AcceptsReturn = $true
    $tbQuery.AcceptsTab = $true
    $tbQuery.ScrollBars = 'Both'
    $tbQuery.WordWrap = $false
    $tbQuery.Font = New-Object System.Drawing.Font('Consolas', 10)
    if ($Preset -and $Preset.Query) { $tbQuery.Text = [string]$Preset.Query } else { $tbQuery.Text = $script:LastQuery }
    $f.Controls.Add($tbQuery)

    $lblNote = New-Control 'Label' 15 612 450 20 'Passwords are never saved.'
    $lblNote.ForeColor = [System.Drawing.Color]::DimGray
    $f.Controls.Add($lblNote)
    $btnOk = New-Control 'Button' 570 608 95 30 'Get plan'
    $btnCancel = New-Control 'Button' 670 608 95 30 'Cancel'
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $f.Controls.Add($btnOk)
    $f.Controls.Add($btnCancel)
    $f.CancelButton = $btnCancel

    $tbUser.Enabled = ($cbAuth.SelectedIndex -eq 1)
    $tbPass.Enabled = ($cbAuth.SelectedIndex -eq 1)
    $cbAuth.Add_SelectedIndexChanged({
        $on = ($cbAuth.SelectedIndex -eq 1)
        $tbUser.Enabled = $on
        $tbPass.Enabled = $on
    })
    $rbAct.Add_CheckedChanged({ $chkRb.Enabled = $rbAct.Checked })
    $tbQuery.Add_KeyDown({
        param($sender, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) { $sender.SelectAll(); $e.SuppressKeyPress = $true }
    })
    $btnOk.Add_Click({
        if (-not $tbServer.Text.Trim()) {
            [void][System.Windows.Forms.MessageBox]::Show($f, 'Enter a server name.', 'SQL Plan Viewer')
            return
        }
        if (-not $tbQuery.Text.Trim()) {
            [void][System.Windows.Forms.MessageBox]::Show($f, 'Enter a query.', 'SQL Plan Viewer')
            return
        }
        $f.DialogResult = [System.Windows.Forms.DialogResult]::OK
    })

    $result = $null
    if ($f.ShowDialog($UI.Form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $result = @{
            Server   = $tbServer.Text.Trim()
            Database = $tbDb.Text.Trim()
            Windows  = ($cbAuth.SelectedIndex -eq 0)
            User     = $tbUser.Text
            Password = $tbPass.Text
            Trust    = $chkTrust.Checked
            Actual   = $rbAct.Checked
            Rollback = $chkRb.Checked
            Timeout  = [int]$nud.Value
            Query    = $tbQuery.Text
        }
        $script:LastQuery = $result.Query
        Save-Settings $result
    }
    $f.Dispose()
    return $result
}

function Get-PlansFromServer {
    param($Cfg)

    $plans = New-Object System.Collections.Generic.List[string]
    $conn = New-SqlConnection $Cfg $Cfg.Database
    try {
        Wait-Task ($conn.OpenAsync())
        $cmd = $conn.CreateCommand()
        $cmd.CommandTimeout = [int]$Cfg.Timeout

        if ($Cfg.Actual -and $Cfg.Rollback) {
            $cmd.CommandText = 'BEGIN TRANSACTION'
            [void]$cmd.ExecuteNonQuery()
        }
        if ($Cfg.Actual) { $cmd.CommandText = 'SET STATISTICS XML ON' } else { $cmd.CommandText = 'SET SHOWPLAN_XML ON' }
        [void]$cmd.ExecuteNonQuery()

        foreach ($batch in [regex]::Split($Cfg.Query, '(?im)^[ \t]*GO[ \t]*\r?$')) {
            if (-not $batch.Trim()) { continue }
            $cmd.CommandText = $batch
            $t = $cmd.ExecuteReaderAsync()
            Wait-Task $t
            $rd = $t.Result
            try {
                do {
                    $isPlan = ($rd.FieldCount -eq 1 -and $rd.GetName(0) -like 'Microsoft SQL Server*Showplan*')
                    while ($true) {
                        $r = $rd.ReadAsync()
                        Wait-Task $r
                        if (-not $r.Result) { break }
                        if ($isPlan) { $plans.Add((Get-ReaderXml $rd)) }
                    }
                    $nx = $rd.NextResultAsync()
                    Wait-Task $nx
                } while ($nx.Result)
            } finally { $rd.Dispose() }
        }
    } finally {
        try {
            if ($Cfg.Actual -and $Cfg.Rollback -and $conn.State -eq [System.Data.ConnectionState]::Open) {
                $c2 = $conn.CreateCommand()
                $c2.CommandText = 'IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION'
                [void]$c2.ExecuteNonQuery()
            }
        } catch { }
        $conn.Dispose()
    }
    return , $plans.ToArray()
}

function Start-QueryPlan {
    param($Preset)
    $cfg = Show-QueryDialog -Preset $Preset
    if (-not $cfg) { return }
    if ($cfg.Actual) {
        $warn = 'This will EXECUTE the query on {0} (database: {1}).' -f $cfg.Server, $cfg.Database
        if ($cfg.Rollback) { $warn += "`r`n`r`nChanges are rolled back afterwards, but side effects outside the transaction (emails, external calls, identity/sequence values, linked servers) are not undone." }
        else { $warn += "`r`n`r`nRollback is OFF - any changes the query makes will be kept." }
        $ans = [System.Windows.Forms.MessageBox]::Show($UI.Form, $warn + "`r`n`r`nContinue?", 'Execute query?', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }
    $kind = 'estimated'
    if ($cfg.Actual) { $kind = 'actual' }
    $UI.Form.Enabled = $false
    $UI.Form.UseWaitCursor = $true
    $UI.Status.Text = 'Running query on {0} ...' -f $cfg.Server
    try {
        $texts = Get-PlansFromServer $cfg
        $UI.Form.Enabled = $true
        $UI.Form.UseWaitCursor = $false
        if (@($texts).Count -eq 0) {
            Show-Info 'The server returned no plans. Is the query empty or comment-only?'
            $UI.Status.Text = 'No plans returned.'
        } else {
            Open-PlanText -Texts ([string[]]$texts) -Source ('{0} ({1} plan)' -f $cfg.Server, $kind)
        }
    } catch {
        Show-Error $_
        $UI.Status.Text = 'Failed.'
    } finally {
        $UI.Form.Enabled = $true
        $UI.Form.UseWaitCursor = $false
        $UI.Form.Activate()
    }
}

# ---------- SQL instance browser ----------

$script:InstancesPath = Join-Path $env:APPDATA 'SqlPlanViewer\instances.json'
$script:Instances = New-Object System.Collections.Generic.List[object]
$script:Pw = @{}              # session-only password cache (never written to disk)
$script:BrowserSel = $null
$script:BrowserSplitSet = $false

function Get-InstancePassword {
    param($Inst)
    return $script:Pw[([string]$Inst.Server).ToLower()]
}

function Get-InstanceCfg {
    param($Inst)
    return @{
        Server   = $Inst.Server
        Windows  = [bool]$Inst.Windows
        User     = $Inst.User
        Trust    = [bool]$Inst.Trust
        Password = (Get-InstancePassword $Inst)
    }
}

function New-SqlConnection {
    param($Cfg, [string]$Database)
    $csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    # Real connection-string keywords: PowerShell exposes this builder as a dictionary.
    $csb['Data Source'] = $Cfg.Server
    if ($Database) { $csb['Initial Catalog'] = $Database }
    if ($Cfg.Windows) { $csb['Integrated Security'] = $true } else { $csb['User ID'] = $Cfg.User; $csb['Password'] = $Cfg.Password }
    $csb['Encrypt'] = $true
    $csb['TrustServerCertificate'] = [bool]$Cfg.Trust
    $csb['Application Name'] = 'SqlPlanViewer'
    $csb['Connect Timeout'] = 15
    $csb['Pooling'] = $false
    return (New-Object System.Data.SqlClient.SqlConnection($csb.ConnectionString))
}

function Invoke-SqlRows {
    param($Inst, [string]$Database, [string]$Query, [hashtable]$Params)
    $conn = New-SqlConnection (Get-InstanceCfg $Inst) $Database
    try {
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Query
        $cmd.CommandTimeout = 30
        if ($Params) { foreach ($k in $Params.Keys) { [void]$cmd.Parameters.AddWithValue($k, $Params[$k]) } }
        $rows = New-Object System.Collections.Generic.List[object]
        $rd = $cmd.ExecuteReader()
        try {
            $cnt = $rd.FieldCount
            while ($rd.Read()) {
                $vals = New-Object 'object[]' $cnt
                [void]$rd.GetValues($vals)
                $rows.Add($vals)
            }
        } finally { $rd.Dispose() }
        return , $rows
    } finally { $conn.Dispose() }
}

function Import-Instances {
    try {
        if (Test-Path -LiteralPath $script:InstancesPath) {
            $j = Get-Content -LiteralPath $script:InstancesPath -Raw | ConvertFrom-Json
            foreach ($i in @($j)) {
                if ($i -and $i.Server) {
                    $script:Instances.Add(@{ Server = [string]$i.Server; Windows = [bool]$i.Windows; User = [string]$i.User; Trust = [bool]$i.Trust })
                }
            }
        }
    } catch { }
}

function Save-Instances {
    try {
        $dir = Split-Path -Parent $script:InstancesPath
        if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir) }
        $list = @(foreach ($i in $script:Instances) { [pscustomobject]@{ Server = $i.Server; Windows = $i.Windows; User = $i.User; Trust = $i.Trust } })
        if ($list.Count -eq 0) { Set-Content -LiteralPath $script:InstancesPath -Value '[]' -Encoding UTF8 }
        else { ConvertTo-Json -InputObject $list | Set-Content -LiteralPath $script:InstancesPath -Encoding UTF8 }
    } catch { }
}

function Show-InstanceDialog {
    param($Existing)
    $f = New-Object System.Windows.Forms.Form
    if ($Existing) { $f.Text = 'Edit SQL Server instance' } else { $f.Text = 'Add SQL Server instance' }
    $f.ClientSize = New-Object System.Drawing.Size(470, 250)
    $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'FixedDialog'
    $f.MaximizeBox = $false
    $f.MinimizeBox = $false
    $f.ShowInTaskbar = $false
    $f.Font = $UI.Font

    $f.Controls.Add((New-Control 'Label' 15 18 75 20 'Instance'))
    $tbServer = New-Control 'TextBox' 95 15 360 23
    $f.Controls.Add($tbServer)
    $f.Controls.Add((New-Control 'Label' 15 51 75 20 'Auth'))
    $cbAuth = New-Control 'ComboBox' 95 48 360 23
    $cbAuth.DropDownStyle = 'DropDownList'
    [void]$cbAuth.Items.Add('Windows authentication')
    [void]$cbAuth.Items.Add('SQL Server authentication')
    $f.Controls.Add($cbAuth)
    $f.Controls.Add((New-Control 'Label' 15 84 75 20 'Login'))
    $tbUser = New-Control 'TextBox' 95 81 360 23
    $f.Controls.Add($tbUser)
    $f.Controls.Add((New-Control 'Label' 15 117 75 20 'Password'))
    $tbPass = New-Control 'TextBox' 95 114 360 23
    $tbPass.UseSystemPasswordChar = $true
    $f.Controls.Add($tbPass)
    $chkTrust = New-Control 'CheckBox' 95 146 360 22 'Trust server certificate (no validation)'
    $f.Controls.Add($chkTrust)
    $note = New-Control 'Label' 95 174 360 34 'The instance list is saved; passwords are kept in memory for this session only.'
    $note.ForeColor = [System.Drawing.Color]::DimGray
    $f.Controls.Add($note)
    $btnOk = New-Control 'Button' 265 210 90 30 'OK'
    $btnCancel = New-Control 'Button' 365 210 90 30 'Cancel'
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $f.Controls.Add($btnOk)
    $f.Controls.Add($btnCancel)
    $f.CancelButton = $btnCancel

    if ($Existing) {
        $tbServer.Text = $Existing.Server
        if ($Existing.Windows) { $cbAuth.SelectedIndex = 0 } else { $cbAuth.SelectedIndex = 1 }
        $tbUser.Text = $Existing.User
        $chkTrust.Checked = [bool]$Existing.Trust
        $pw = Get-InstancePassword $Existing
        if ($pw) { $tbPass.Text = [string]$pw }
    } else {
        $cbAuth.SelectedIndex = 0
        $chkTrust.Checked = $false
    }
    $tbUser.Enabled = ($cbAuth.SelectedIndex -eq 1)
    $tbPass.Enabled = ($cbAuth.SelectedIndex -eq 1)
    $cbAuth.Add_SelectedIndexChanged({
        $on = ($cbAuth.SelectedIndex -eq 1)
        $tbUser.Enabled = $on
        $tbPass.Enabled = $on
    })
    $btnOk.Add_Click({
        if (-not $tbServer.Text.Trim()) {
            [void][System.Windows.Forms.MessageBox]::Show($f, 'Enter an instance name.', 'SQL Plan Viewer')
            return
        }
        $f.DialogResult = [System.Windows.Forms.DialogResult]::OK
    })

    $result = $null
    if ($f.ShowDialog($UI.Form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $result = @{
            Server   = $tbServer.Text.Trim()
            Windows  = ($cbAuth.SelectedIndex -eq 0)
            User     = $tbUser.Text
            Trust    = $chkTrust.Checked
            Password = $tbPass.Text
        }
    }
    $f.Dispose()
    return $result
}

function Update-Instance {
    param($Inst, $R)
    $old = ([string]$Inst.Server).ToLower()
    $Inst.Server = $R.Server
    $Inst.Windows = $R.Windows
    $Inst.User = $R.User
    $Inst.Trust = $R.Trust
    if ($old -ne $R.Server.ToLower()) { $script:Pw.Remove($old) }
    if ($R.Windows) { $script:Pw.Remove($R.Server.ToLower()) }
    elseif ($R.Password) { $script:Pw[$R.Server.ToLower()] = $R.Password }
    Save-Instances
}

function Add-DummyChild {
    param($Node)
    $d = New-Object System.Windows.Forms.TreeNode('Loading...')
    $d.Tag = 'dummy'
    [void]$Node.Nodes.Add($d)
}

function Test-DummyNode {
    param($Node)
    return ($Node.Nodes.Count -eq 1 -and ($Node.Nodes[0].Tag -is [string]) -and $Node.Nodes[0].Tag -eq 'dummy')
}

function Add-DatabaseChildren {
    param($DbNode)
    $info = $DbNode.Tag
    foreach ($k in @(@('ViewFolder', 'Views'), @('ProcFolder', 'Stored procedures'))) {
        $tn = New-Object System.Windows.Forms.TreeNode($k[1])
        $tn.Tag = @{ Kind = $k[0]; Instance = $info.Instance; Database = $info.Database }
        Add-DummyChild $tn
        [void]$DbNode.Nodes.Add($tn)
    }
}

function Add-InstanceNode {
    param($Inst)
    $tn = New-Object System.Windows.Forms.TreeNode([string]$Inst.Server)
    $tn.Tag = @{ Kind = 'Instance'; Instance = $Inst }
    Add-DummyChild $tn
    [void]$UI.BTree.Nodes.Add($tn)
    return $tn
}

function Get-RootNode {
    param($Node)
    $n = $Node
    while ($n -and $n.Parent) { $n = $n.Parent }
    return $n
}

function Reset-BrowserNode {
    param($Node)
    $info = $Node.Tag
    if ($info -isnot [hashtable]) { return }
    $Node.Nodes.Clear()
    if ($info.Kind -eq 'Database') { Add-DatabaseChildren $Node } else { Add-DummyChild $Node }
    $Node.Collapse()
}

function Open-BrowserNode {
    # Lazy-loads the children of a tree node. Returns $false if the node should not expand.
    param($Node)
    $info = $Node.Tag
    $inst = $info.Instance
    if ($info.Kind -eq 'Instance' -and -not $inst.Windows -and -not (Get-InstancePassword $inst)) {
        $r = Show-InstanceDialog $inst
        if (-not $r) { return $false }
        Update-Instance $inst $r
    }
    [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $newNodes = New-Object System.Collections.Generic.List[object]
        if ($info.Kind -eq 'Instance') {
            $rows = Invoke-SqlRows $inst '' 'SELECT name FROM sys.databases WHERE state = 0 AND HAS_DBACCESS(name) = 1 ORDER BY CASE WHEN database_id <= 4 THEN 1 ELSE 0 END, name'
            foreach ($row in $rows) {
                $tn = New-Object System.Windows.Forms.TreeNode([string]$row[0])
                $tn.Tag = @{ Kind = 'Database'; Instance = $inst; Database = [string]$row[0] }
                Add-DatabaseChildren $tn
                $newNodes.Add($tn)
            }
        } else {
            $type = 'V'; $kindName = 'View'
            if ($info.Kind -eq 'ProcFolder') { $type = 'P'; $kindName = 'Proc' }
            $sys = ''
            if (-not $UI.BrowserSys.Checked) { $sys = ' AND o.is_ms_shipped = 0' }
            $sql = "SELECT s.name, o.name FROM sys.objects o JOIN sys.schemas s ON s.schema_id = o.schema_id WHERE o.type = '$type'$sys ORDER BY s.name, o.name"
            $rows = Invoke-SqlRows $inst $info.Database $sql
            foreach ($row in $rows) {
                $tn = New-Object System.Windows.Forms.TreeNode(('{0}.{1}' -f $row[0], $row[1]))
                $tn.Tag = @{ Kind = $kindName; Instance = $inst; Database = $info.Database; Schema = [string]$row[0]; Name = [string]$row[1] }
                $newNodes.Add($tn)
            }
        }
        $Node.Nodes.Clear()
        foreach ($tn in $newNodes) { [void]$Node.Nodes.Add($tn) }
        if ($newNodes.Count -eq 0) {
            $e = New-Object System.Windows.Forms.TreeNode('(none)')
            $e.Tag = 'empty'
            $e.ForeColor = [System.Drawing.Color]::Gray
            [void]$Node.Nodes.Add($e)
        }
        return $true
    } catch {
        Show-Error $_
        return $false
    } finally {
        [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
    }
}

function Get-QuotedName {
    param($Info)
    return ('[{0}].[{1}]' -f (([string]$Info.Schema) -replace '\]', ']]'), (([string]$Info.Name) -replace '\]', ']]'))
}

function Show-BrowserSelection {
    param($Node)
    $script:BrowserSel = $null
    $UI.BtnObjPlan.Enabled = $false
    $UI.BtnCopyDef.Enabled = $false
    $info = $null
    if ($Node) { $info = $Node.Tag }
    if ($info -isnot [hashtable]) { $UI.DefBox.Text = ''; return }
    if ($info.Kind -notin @('View', 'Proc')) {
        $UI.DefBox.Text = 'Expand an instance, a database and then Views or Stored procedures. Select an object to see its definition.'
        return
    }
    $full = Get-QuotedName $info
    [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $rows = Invoke-SqlRows $info.Instance $info.Database 'SELECT m.definition FROM sys.sql_modules m WHERE m.object_id = OBJECT_ID(@n)' @{ '@n' = $full }
        $def = $null
        foreach ($row in $rows) { if ($row[0] -isnot [System.DBNull]) { $def = [string]$row[0] } }
        if ($null -eq $def) { $def = '-- Definition not available (encrypted, or no permission to view it).' }
        $def = $def -replace "(?<!`r)`n", "`r`n"
        $UI.DefBox.Text = ('-- {0} / {1} / {2}' -f $info.Instance.Server, $info.Database, $full) + "`r`n`r`n" + $def
        $script:BrowserSel = $info
        $UI.BtnObjPlan.Enabled = $true
        $UI.BtnCopyDef.Enabled = $true
    } catch {
        $UI.DefBox.Text = '-- ' + $_.Exception.Message
    } finally {
        [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
    }
}

function Start-ObjectPlan {
    $info = $script:BrowserSel
    if (-not $info) { return }
    $full = Get-QuotedName $info
    $inst = $info.Instance
    if ($info.Kind -eq 'View') {
        $q = "SELECT * FROM $full;"
    } else {
        $parts = @(); $skipped = @()
        try {
            $rows = Invoke-SqlRows $inst $info.Database 'SELECT p.name, p.is_output, p.has_default_value FROM sys.parameters p WHERE p.object_id = OBJECT_ID(@n) AND p.parameter_id > 0 ORDER BY p.parameter_id' @{ '@n' = $full }
            foreach ($row in $rows) {
                if ([bool]$row[2]) { $skipped += [string]$row[0] }
                else {
                    $p = '{0} = NULL' -f $row[0]
                    if ([bool]$row[1]) { $p += ' OUTPUT' }
                    $parts += $p
                }
            }
        } catch { Show-Error $_; return }
        $q = "-- Replace the NULLs with representative parameter values.`r`n"
        if ($skipped.Count -gt 0) { $q += "-- Parameters with defaults (add them if needed): $($skipped -join ', ')`r`n" }
        $q += "EXEC $full"
        if ($parts.Count -gt 0) { $q += ' ' + ($parts -join ', ') }
        $q += ';'
    }
    Start-QueryPlan -Preset @{
        Server   = $inst.Server
        Database = $info.Database
        Windows  = [bool]$inst.Windows
        User     = $inst.User
        Trust    = [bool]$inst.Trust
        Password = (Get-InstancePassword $inst)
        Actual   = $false
        Query    = $q
    }
}

function Invoke-AddInstance {
    $r = Show-InstanceDialog $null
    if (-not $r) { return }
    foreach ($tn in $UI.BTree.Nodes) {
        if ($tn.Tag.Instance.Server -ieq $r.Server) {
            $UI.BTree.SelectedNode = $tn
            Show-Info 'That instance is already in the list.'
            return
        }
    }
    $inst = @{ Server = $r.Server; Windows = $r.Windows; User = $r.User; Trust = $r.Trust }
    $script:Instances.Add($inst)
    Update-Instance $inst $r
    $UI.BTree.SelectedNode = (Add-InstanceNode $inst)
}

function Invoke-EditInstance {
    $root = Get-RootNode $UI.BTree.SelectedNode
    if (-not $root) { return }
    $inst = $root.Tag.Instance
    $r = Show-InstanceDialog $inst
    if (-not $r) { return }
    Update-Instance $inst $r
    $root.Text = $inst.Server
    Reset-BrowserNode $root
}

function Invoke-RemoveInstance {
    $root = Get-RootNode $UI.BTree.SelectedNode
    if (-not $root) { return }
    $inst = $root.Tag.Instance
    $ans = [System.Windows.Forms.MessageBox]::Show($UI.Form, ('Remove {0} from the list?' -f $inst.Server), 'SQL Plan Viewer', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    [void]$script:Instances.Remove($inst)
    $script:Pw.Remove(([string]$inst.Server).ToLower())
    Save-Instances
    $root.Remove()
}

function Invoke-RefreshBrowser {
    $n = $UI.BTree.SelectedNode
    if (-not $n) { return }
    # refresh the nearest node that owns lazily loaded children
    while ($n -and ($n.Tag -isnot [hashtable] -or $n.Tag.Kind -in @('View', 'Proc'))) { $n = $n.Parent }
    if ($n) { Reset-BrowserNode $n }
}

function Reset-AllInstances {
    foreach ($n in @($UI.BTree.Nodes)) { Reset-BrowserNode $n }
    Show-BrowserSelection $null
}

function Add-StripButton {
    param($Strip, [string]$Text, [scriptblock]$Action, [bool]$CheckOnClick = $false)
    $b = New-Object System.Windows.Forms.ToolStripButton($Text)
    $b.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::Text
    $b.CheckOnClick = $CheckOnClick
    $b.Add_Click($Action)
    [void]$Strip.Items.Add($b)
    return $b
}

# ---------- loading plans ----------

function Open-PlanText {
    param([string[]]$Texts, [string]$Source)
    $all = New-Object System.Collections.Generic.List[object]
    $i = 0
    try {
        foreach ($t in $Texts) {
            foreach ($st in @(Read-PlanXml $t)) {
                $st.SourceIndex = $i
                $st.Index = $all.Count + 1
                $all.Add($st)
            }
            $i++
        }
    } catch { Show-Error $_; return }

    if ($all.Count -eq 0) {
        Show-Info 'No statements with a query plan were found in that XML.'
        return
    }
    $total = 0.0
    foreach ($st in $all) { $total += $st.Cost }
    if ($total -le 0) { $total = 1.0 }
    foreach ($st in $all) { $st.CostPct = $st.Cost / $total * 100.0 }

    $script:Statements = $all
    $script:PlanTexts = $Texts
    $UI.Form.Text = 'SQL Plan Viewer - ' + $Source

    $UI.Stmts.BeginUpdate()
    $UI.Stmts.Items.Clear()
    foreach ($st in $all) {
        $it = New-Object System.Windows.Forms.ListViewItem([string]$st.Index)
        [void]$it.SubItems.Add($st.Type)
        [void]$it.SubItems.Add(('{0:0.#}%' -f $st.CostPct))
        [void]$it.SubItems.Add(('{0:0.####}' -f $st.Cost))
        $one = $st.OneLine
        if ($one.Length -gt 300) { $one = $one.Substring(0, 300) + '...' }
        [void]$it.SubItems.Add($one)
        $it.Tag = $st
        [void]$UI.Stmts.Items.Add($it)
    }
    $UI.Stmts.EndUpdate()
    $UI.Stmts.Items[0].Selected = $true
    $tabs.SelectedTab = $tabPlan
}

function Open-PlanFile {
    param([string]$Path)
    try {
        if ((Get-Item -LiteralPath $Path).Length -gt 100MB) { throw 'File is larger than 100 MB - refusing to load it.' }
        $text = [System.IO.File]::ReadAllText($Path)
        Open-PlanText -Texts @($text) -Source ([System.IO.Path]::GetFileName($Path))
    } catch { Show-Error $_ }
}

function Invoke-OpenFile {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'SQL Server execution plans (*.sqlplan;*.xml)|*.sqlplan;*.xml|All files (*.*)|*.*'
    if ($dlg.ShowDialog($UI.Form) -eq [System.Windows.Forms.DialogResult]::OK) { Open-PlanFile $dlg.FileName }
}

function Invoke-PasteXml {
    $text = ''
    try { $text = [System.Windows.Forms.Clipboard]::GetText() } catch { }
    if (-not $text -or -not $text.TrimStart().StartsWith('<')) {
        Show-Info 'The clipboard does not contain plan XML. Copy the XML (or "Show Execution Plan XML" from SSMS) first.'
        return
    }
    Open-PlanText -Texts @($text) -Source 'clipboard'
}

function Save-PlanFile {
    $st = $script:CurrentStmt
    if (-not $st) { return }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'SQL Server execution plan (*.sqlplan)|*.sqlplan'
    $dlg.FileName = 'plan.sqlplan'
    if ($dlg.ShowDialog($UI.Form) -eq [System.Windows.Forms.DialogResult]::OK) {
        try { [System.IO.File]::WriteAllText($dlg.FileName, $script:PlanTexts[$st.SourceIndex], [System.Text.Encoding]::Unicode) } catch { Show-Error $_ }
    }
}

# ---------- drawing ----------

function Invoke-PlanDraw {
    param([System.Drawing.Graphics]$G, $St)
    $G.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $G.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $W = $script:NodeW
    $H = $script:NodeH

    # connectors (arrows point towards the consumer, like SSMS)
    $edgeColor = [System.Drawing.Color]::FromArgb(150, 150, 160)
    foreach ($n in $St.Nodes) {
        foreach ($c in $n.Children) {
            $rows = $c.EstRows
            if ($c.HasRuntime) { $rows = $c.ActualRows }
            $pw = [single][math]::Min(7.0, [math]::Max(1.2, [math]::Log10([math]::Max($rows, 1.0) + 1.0) * 1.2))
            $pen = New-Object System.Drawing.Pen($edgeColor, $pw)
            $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::ArrowAnchor
            $x1 = [single]$c.X
            $y1 = [single]($c.Y + $H / 2)
            $x2 = [single]($n.X + $W)
            $y2 = [single]($n.Y + $H / 2)
            $mx = [single](($x1 + $x2) / 2)
            $pts = [System.Drawing.PointF[]]@(
                (New-Object System.Drawing.PointF($x1, $y1)),
                (New-Object System.Drawing.PointF($mx, $y1)),
                (New-Object System.Drawing.PointF($mx, $y2)),
                (New-Object System.Drawing.PointF($x2, $y2)))
            $G.DrawLines($pen, $pts)
            $pen.Dispose()
        }
    }

    foreach ($n in $St.Nodes) {
        $t = [math]::Min(1.0, $n.CostPct / 50.0)
        $fill = [System.Drawing.Color]::FromArgb(255, [int](244 + 11 * $t), [int](247 - 100 * $t), [int](252 - 140 * $t))
        $path = New-RoundRectPath $n.X $n.Y $W $H 8
        $brush = New-Object System.Drawing.SolidBrush($fill)
        $G.FillPath($brush, $path)
        $brush.Dispose()
        if ([object]::ReferenceEquals($n, $script:SelNode)) {
            $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(0, 120, 215), [single]3)
        } else {
            $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(140, 140, 150), [single]1)
        }
        $G.DrawPath($pen, $path)
        $pen.Dispose()
        $path.Dispose()

        $hasWarn = ($n.Warnings.Count -gt 0)
        $titleW = $W - 16
        if ($hasWarn) { $titleW = $W - 36 }
        $r1 = New-Object System.Drawing.RectangleF(($n.X + 8), ($n.Y + 5), $titleW, 17)
        $G.DrawString($n.Display, $script:FontBold, [System.Drawing.Brushes]::Black, $r1, $script:Sf)

        $sub = $n.ObjectShort
        if (-not $sub -and $n.LogicalOp -and $n.LogicalOp -ne $n.PhysicalOp) { $sub = $n.LogicalOp }
        $r2 = New-Object System.Drawing.RectangleF(($n.X + 8), ($n.Y + 23), ($W - 16), 16)
        $G.DrawString($sub, $script:FontNorm, [System.Drawing.Brushes]::DimGray, $r2, $script:Sf)

        $costTxt = 'Cost: {0:0.#}%' -f $n.CostPct
        if ($n.HasRuntime) { $costTxt += '   {0:N0} ms' -f $n.ActualElapsed }
        $r3 = New-Object System.Drawing.RectangleF(($n.X + 8), ($n.Y + 40), ($W - 16), 16)
        $G.DrawString($costTxt, $script:FontNorm, [System.Drawing.Brushes]::Black, $r3, $script:Sf)

        if ($n.HasRuntime) { $rowTxt = 'Rows: {0} of {1}' -f (Format-Count $n.ActualRows), (Format-Count ($n.EstRows * [math]::Max($n.ActualExec, 1))) }
        else { $rowTxt = 'Est. rows: {0}' -f (Format-Count $n.EstRows) }
        $r4 = New-Object System.Drawing.RectangleF(($n.X + 8), ($n.Y + 57), ($W - 16), 16)
        $G.DrawString($rowTxt, $script:FontNorm, [System.Drawing.Brushes]::Black, $r4, $script:Sf)

        if ($hasWarn) {
            $rb = New-Object System.Drawing.RectangleF(($n.X + $W - 26), ($n.Y + 5), 18, 18)
            $G.FillEllipse([System.Drawing.Brushes]::Firebrick, $rb)
            $G.DrawString('!', $script:FontBold, [System.Drawing.Brushes]::White, $rb, $script:SfCenter)
        }
    }
}

function Update-CanvasSize {
    $st = $script:CurrentStmt
    if ($st) {
        $UI.Canvas.AutoScrollMinSize = New-Object System.Drawing.Size([int]($st.MaxX * $script:Zoom), [int]($st.MaxY * $script:Zoom))
    } else {
        $UI.Canvas.AutoScrollMinSize = New-Object System.Drawing.Size(0, 0)
    }
}

function Set-Zoom {
    param([double]$Z)
    $script:Zoom = [math]::Max(0.2, [math]::Min(2.5, $Z))
    Update-CanvasSize
    $UI.ZoomLabel.Text = ' {0:0}% ' -f ($script:Zoom * 100)
    $UI.Canvas.Invalidate()
}

function Set-ZoomFit {
    $st = $script:CurrentStmt
    if (-not $st -or $st.MaxX -le 0 -or $st.MaxY -le 0) { return }
    $cw = [math]::Max(100, $UI.Canvas.ClientSize.Width)
    $ch = [math]::Max(100, $UI.Canvas.ClientSize.Height)
    Set-Zoom ([math]::Min(1.0, [math]::Min($cw / $st.MaxX, $ch / $st.MaxY)))
}

function Get-NodeAt {
    param([int]$PX, [int]$PY)
    $st = $script:CurrentStmt
    if ($null -eq $st) { return $null }
    $asp = $UI.Canvas.AutoScrollPosition
    $lx = ($PX - $asp.X) / $script:Zoom
    $ly = ($PY - $asp.Y) / $script:Zoom
    foreach ($n in $st.Nodes) {
        if ($lx -ge $n.X -and $lx -le ($n.X + $script:NodeW) -and $ly -ge $n.Y -and $ly -le ($n.Y + $script:NodeH)) { return $n }
    }
    return $null
}

function Get-NodeTip {
    param($N)
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add($N.Display)
    if ($N.ObjectFull) { $lines.Add($N.ObjectFull) }
    $lines.Add(('Cost: {0:0.#}% (operator {1:0.000000}, subtree {2:0.000000})' -f $N.CostPct, $N.SelfCost, $N.SubtreeCost))
    if ($N.HasRuntime) {
        $lines.Add(('Actual rows: {0:N0}   executions: {1:N0}' -f $N.ActualRows, $N.ActualExec))
        $lines.Add(('Elapsed: {0:N0} ms   CPU: {1:N0} ms' -f $N.ActualElapsed, $N.ActualCpu))
    }
    $lines.Add('Estimated rows: ' + (Format-Count $N.EstRows))
    foreach ($w in $N.Warnings) { $lines.Add('(!) ' + $w) }
    return ($lines -join "`n")
}

function Export-PlanImage {
    $st = $script:CurrentStmt
    if (-not $st) { return }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'PNG image (*.png)|*.png'
    $dlg.FileName = 'plan.png'
    if ($dlg.ShowDialog($UI.Form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
    try {
        $bmp = New-Object System.Drawing.Bitmap([int]$st.MaxX, [int]$st.MaxY)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        try {
            $g.Clear([System.Drawing.Color]::White)
            Invoke-PlanDraw $g $st
            $bmp.Save($dlg.FileName, [System.Drawing.Imaging.ImageFormat]::Png)
        } finally { $g.Dispose(); $bmp.Dispose() }
    } catch { Show-Error $_ }
}

# ---------- populating the panes ----------

function Add-TreeNode {
    param($N, $Collection)
    $obj = ''
    if ($N.ObjectShort) { $obj = ' - ' + $N.ObjectShort }
    $label = '{0}{1}   [{2:0.#}%]' -f $N.Display, $obj, $N.CostPct
    if ($N.Warnings.Count -gt 0) { $label += '  (!)' }
    $tn = New-Object System.Windows.Forms.TreeNode($label)
    $tn.Tag = $N
    if ($N.CostPct -ge 25) { $tn.ForeColor = [System.Drawing.Color]::Firebrick }
    $N.TreeNode = $tn
    [void]$Collection.Add($tn)
    foreach ($c in $N.Children) { Add-TreeNode $c $tn.Nodes }
}

function Show-NodeDetails {
    param($N)
    $UI.Props.BeginUpdate()
    $UI.Props.Items.Clear()
    $UI.Props.Groups.Clear()
    $UI.ValueBox.Text = ''
    if ($N) {
        $groups = @{}
        foreach ($d in (Get-NodeDetails $N)) {
            if (-not $groups.ContainsKey($d.G)) {
                $grp = New-Object System.Windows.Forms.ListViewGroup($d.G, $d.G)
                [void]$UI.Props.Groups.Add($grp)
                $groups[$d.G] = $grp
            }
            $it = New-Object System.Windows.Forms.ListViewItem([string]$d.N)
            [void]$it.SubItems.Add($d.V)
            $it.Group = $groups[$d.G]
            [void]$UI.Props.Items.Add($it)
        }
    }
    $UI.Props.EndUpdate()
}

function Select-PlanNode {
    param($N, [switch]$Scroll)
    if ($script:Syncing) { return }
    $script:Syncing = $true
    try {
        $script:SelNode = $N
        if ($N -and $N.TreeNode -and $UI.Tree.SelectedNode -ne $N.TreeNode) { $UI.Tree.SelectedNode = $N.TreeNode }
        Show-NodeDetails $N
        if ($N -and $Scroll) {
            $x = [int][math]::Max(0, $N.X * $script:Zoom - ($UI.Canvas.ClientSize.Width - $script:NodeW * $script:Zoom) / 2)
            $y = [int][math]::Max(0, $N.Y * $script:Zoom - ($UI.Canvas.ClientSize.Height - $script:NodeH * $script:Zoom) / 2)
            $UI.Canvas.AutoScrollPosition = New-Object System.Drawing.Point($x, $y)
        }
        $UI.Canvas.Invalidate()
    } finally { $script:Syncing = $false }
}

function Select-Statement {
    param($St)
    $script:CurrentStmt = $St
    $script:SelNode = $null

    $script:LayoutY = [double]$script:Margin
    Set-PlanLayout $St.Root 0
    $maxX = 0.0
    foreach ($n in $St.Nodes) { $maxX = [math]::Max($maxX, $n.X + $script:NodeW) }
    $St.MaxX = $maxX + $script:Margin
    $St.MaxY = $script:LayoutY - $script:GapY + $script:Margin

    # tree
    $UI.Tree.BeginUpdate()
    $UI.Tree.Nodes.Clear()
    Add-TreeNode $St.Root $UI.Tree.Nodes
    $UI.Tree.ExpandAll()
    $UI.Tree.EndUpdate()

    # top operators
    $UI.Top.BeginUpdate()
    $UI.Top.Items.Clear()
    foreach ($n in @($St.Nodes | Sort-Object SelfCost -Descending | Select-Object -First 20)) {
        $it = New-Object System.Windows.Forms.ListViewItem($n.Display)
        [void]$it.SubItems.Add($n.ObjectShort)
        [void]$it.SubItems.Add(('{0:0.#}%' -f $n.CostPct))
        $rows = Format-Count $n.EstRows
        if ($n.HasRuntime) { $rows = Format-Count $n.ActualRows }
        [void]$it.SubItems.Add($rows)
        $ms = ''
        if ($n.HasRuntime) { $ms = '{0:N0}' -f $n.ActualElapsed }
        [void]$it.SubItems.Add($ms)
        $it.Tag = $n
        [void]$UI.Top.Items.Add($it)
    }
    $UI.Top.EndUpdate()

    # findings
    $findings = Get-Findings $St
    $UI.Find.BeginUpdate()
    $UI.Find.Items.Clear()
    foreach ($f in $findings) {
        $it = New-Object System.Windows.Forms.ListViewItem($f.Severity)
        [void]$it.SubItems.Add($f.Operator)
        [void]$it.SubItems.Add($f.Message)
        if ($f.Severity -eq 'Warning') { $it.ForeColor = [System.Drawing.Color]::Firebrick }
        $it.Tag = $f
        [void]$UI.Find.Items.Add($it)
    }
    $UI.Find.EndUpdate()
    $nWarn = @($findings | Where-Object { $_.Severity -eq 'Warning' }).Count
    $UI.TabFind.Text = 'Warnings && findings ({0})' -f $findings.Count
    if ($nWarn -gt 0) { $UI.TabFind.Text = 'Warnings && findings ({0}, {1} warn)' -f $findings.Count, $nWarn }

    # missing indexes
    $UI.Miss.Items.Clear()
    $UI.MissScript.Text = ''
    foreach ($mi in $St.MissingIndexes) {
        $it = New-Object System.Windows.Forms.ListViewItem(('{0:0.#}%' -f $mi.Impact))
        [void]$it.SubItems.Add($mi.Table)
        [void]$it.SubItems.Add($mi.Key)
        [void]$it.SubItems.Add($mi.Include)
        $it.Tag = $mi
        [void]$UI.Miss.Items.Add($it)
    }
    $UI.TabMiss.Text = 'Missing indexes ({0})' -f $St.MissingIndexes.Count
    if ($UI.Miss.Items.Count -gt 0) { $UI.Miss.Items[0].Selected = $true }

    # summary
    $UI.Summary.Text = Get-StatementSummary $St

    Update-CanvasSize
    $UI.Canvas.AutoScrollPosition = New-Object System.Drawing.Point(0, 0)
    $UI.Status.Text = 'Statement {0} of {1}   |   {2} operators   |   estimated cost {3:0.####}' -f $St.Index, $script:Statements.Count, $St.Nodes.Count, $St.Cost
    Select-PlanNode $St.Root
}

#endregion

#region Main window

$form = New-Object System.Windows.Forms.Form
$form.Text = 'SQL Plan Viewer'
$form.Size = New-Object System.Drawing.Size(1380, 840)
$form.StartPosition = 'CenterScreen'
$form.KeyPreview = $true
$form.AllowDrop = $true
$UI.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.Font = $UI.Font
$UI.Form = $form
try {
    $icoFile = $null
    if ($PSScriptRoot) { $icoFile = Join-Path $PSScriptRoot 'SqlPlanViewer.ico' }
    if ($icoFile -and (Test-Path -LiteralPath $icoFile)) {
        $form.Icon = New-Object System.Drawing.Icon($icoFile)
    } else {
        $exePath = (Get-Process -Id $PID).Path     # compiled .exe: use the icon embedded in it
        if ($exePath -and $exePath -notmatch '(powershell|pwsh)(_ise)?\.exe$') { $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($exePath) }
    }
} catch { }
$UI.Tip = New-Object System.Windows.Forms.ToolTip
$UI.Tip.InitialDelay = 300
$UI.Tip.AutoPopDelay = 20000

# layout containers (Fill control first, then bottom/top docked bars)
$main = New-Object System.Windows.Forms.SplitContainer
$main.Dock = 'Fill'
$left = New-Object System.Windows.Forms.SplitContainer
$left.Dock = 'Fill'
$left.Orientation = 'Horizontal'
$right = New-Object System.Windows.Forms.SplitContainer
$right.Dock = 'Fill'
$right.Orientation = 'Horizontal'
$main.Panel1.Controls.Add($left)
$main.Panel2.Controls.Add($right)
$form.Controls.Add($main)

$status = New-Object System.Windows.Forms.StatusStrip
$UI.Status = New-Object System.Windows.Forms.ToolStripStatusLabel
$UI.Status.Text = 'Open a .sqlplan file, paste plan XML, or get a plan from SQL Server.'
[void]$status.Items.Add($UI.Status)
$form.Controls.Add($status)

$tool = New-Object System.Windows.Forms.ToolStrip
$tool.GripStyle = 'Hidden'
$tool.Padding = New-Object System.Windows.Forms.Padding(6, 4, 6, 4)
$form.Controls.Add($tool)

function Add-ToolButton {
    param([string]$Text, [scriptblock]$Action, [string]$Tip = '')
    $b = New-Object System.Windows.Forms.ToolStripButton($Text)
    $b.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::Text
    if ($Tip) { $b.ToolTipText = $Tip }
    $b.Add_Click($Action)
    [void]$tool.Items.Add($b)
}
Add-ToolButton 'Open plan...' { Invoke-OpenFile } 'Open a .sqlplan file (Ctrl+O)'
Add-ToolButton 'Paste plan XML' { Invoke-PasteXml } 'Load plan XML from the clipboard (Ctrl+V)'
Add-ToolButton 'Get plan from SQL Server...' { Start-QueryPlan } 'Run a query and capture its plan (Ctrl+R)'
Add-ToolButton 'Save plan...' { Save-PlanFile } 'Save the current plan as .sqlplan'
[void]$tool.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
Add-ToolButton 'Zoom -' { Set-Zoom ($script:Zoom - 0.1) } 'Zoom out (Ctrl+wheel)'
$UI.ZoomLabel = New-Object System.Windows.Forms.ToolStripLabel(' 100% ')
[void]$tool.Items.Add($UI.ZoomLabel)
Add-ToolButton 'Zoom +' { Set-Zoom ($script:Zoom + 0.1) } 'Zoom in (Ctrl+wheel)'
Add-ToolButton 'Fit' { Set-ZoomFit } 'Fit the plan to the window'
Add-ToolButton '100%' { Set-Zoom 1.0 }
[void]$tool.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
Add-ToolButton 'Export PNG...' { Export-PlanImage } 'Save the graphical plan as an image'

# statements list
$UI.Stmts = New-Object System.Windows.Forms.ListView
$UI.Stmts.Dock = 'Fill'
$UI.Stmts.View = 'Details'
$UI.Stmts.FullRowSelect = $true
$UI.Stmts.HideSelection = $false
$UI.Stmts.MultiSelect = $false
$UI.Stmts.GridLines = $true
[void]$UI.Stmts.Columns.Add('#', 35)
[void]$UI.Stmts.Columns.Add('Type', 90)
[void]$UI.Stmts.Columns.Add('Cost %', 60)
[void]$UI.Stmts.Columns.Add('Est. cost', 80)
[void]$UI.Stmts.Columns.Add('Statement', 900)
Set-DoubleBuffered $UI.Stmts
$left.Panel1.Controls.Add($UI.Stmts)

# tabs
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$left.Panel2.Controls.Add($tabs)
function New-Tab {
    param([string]$Title)
    $p = New-Object System.Windows.Forms.TabPage($Title)
    [void]$tabs.TabPages.Add($p)
    return $p
}
$tabPlan = New-Tab 'Graphical plan'
$tabTree = New-Tab 'Operator tree'
$tabTop  = New-Tab 'Top operators'
$UI.TabFind = New-Tab 'Warnings && findings'
$UI.TabMiss = New-Tab 'Missing indexes'
$tabSum  = New-Tab 'Statement && parameters'

# graphical canvas
$canvas = New-Object System.Windows.Forms.Panel
$canvas.Dock = 'Fill'
$canvas.AutoScroll = $true
$canvas.BackColor = [System.Drawing.Color]::White
$canvas.TabStop = $true
Set-DoubleBuffered $canvas
try {
    $setStyle = [System.Windows.Forms.Control].GetMethod('SetStyle', [System.Reflection.BindingFlags]'Instance,NonPublic')
    [void]$setStyle.Invoke($canvas, @([System.Windows.Forms.ControlStyles]::Selectable, $true))
} catch { }
$UI.Canvas = $canvas
$tabPlan.Controls.Add($canvas)

# operator tree
$UI.Tree = New-Object System.Windows.Forms.TreeView
$UI.Tree.Dock = 'Fill'
$UI.Tree.HideSelection = $false
$UI.Tree.ShowLines = $true
Set-DoubleBuffered $UI.Tree
$tabTree.Controls.Add($UI.Tree)

# top operators
$UI.Top = New-Object System.Windows.Forms.ListView
$UI.Top.Dock = 'Fill'
$UI.Top.View = 'Details'
$UI.Top.FullRowSelect = $true
$UI.Top.HideSelection = $false
$UI.Top.GridLines = $true
[void]$UI.Top.Columns.Add('Operator', 230)
[void]$UI.Top.Columns.Add('Object', 320)
[void]$UI.Top.Columns.Add('Cost %', 70)
[void]$UI.Top.Columns.Add('Rows', 90)
[void]$UI.Top.Columns.Add('Elapsed ms', 90)
$tabTop.Controls.Add($UI.Top)

# findings
$UI.Find = New-Object System.Windows.Forms.ListView
$UI.Find.Dock = 'Fill'
$UI.Find.View = 'Details'
$UI.Find.FullRowSelect = $true
$UI.Find.HideSelection = $false
$UI.Find.GridLines = $true
[void]$UI.Find.Columns.Add('Severity', 70)
[void]$UI.Find.Columns.Add('Operator', 230)
[void]$UI.Find.Columns.Add('Message', 900)
$UI.TabFind.Controls.Add($UI.Find)

# missing indexes
$UI.Miss = New-Object System.Windows.Forms.ListView
$UI.Miss.Dock = 'Fill'
$UI.Miss.View = 'Details'
$UI.Miss.FullRowSelect = $true
$UI.Miss.HideSelection = $false
$UI.Miss.GridLines = $true
[void]$UI.Miss.Columns.Add('Impact', 65)
[void]$UI.Miss.Columns.Add('Table', 260)
[void]$UI.Miss.Columns.Add('Key columns', 280)
[void]$UI.Miss.Columns.Add('Include columns', 400)
$UI.MissScript = New-Object System.Windows.Forms.TextBox
$UI.MissScript.Dock = 'Bottom'
$UI.MissScript.Height = 110
$UI.MissScript.Multiline = $true
$UI.MissScript.ReadOnly = $true
$UI.MissScript.ScrollBars = 'Both'
$UI.MissScript.WordWrap = $false
$UI.MissScript.Font = New-Object System.Drawing.Font('Consolas', 10)
$UI.TabMiss.Controls.Add($UI.Miss)
$UI.TabMiss.Controls.Add($UI.MissScript)

# summary
$UI.Summary = New-Object System.Windows.Forms.TextBox
$UI.Summary.Dock = 'Fill'
$UI.Summary.Multiline = $true
$UI.Summary.ReadOnly = $true
$UI.Summary.ScrollBars = 'Both'
$UI.Summary.WordWrap = $false
$UI.Summary.Font = New-Object System.Drawing.Font('Consolas', 10)
$tabSum.Controls.Add($UI.Summary)

# SQL browser tab
$UI.TabBrowser = New-Tab 'SQL browser'
$UI.BSplit = New-Object System.Windows.Forms.SplitContainer
$UI.BSplit.Dock = 'Fill'
$UI.BTree = New-Object System.Windows.Forms.TreeView
$UI.BTree.Dock = 'Fill'
$UI.BTree.HideSelection = $false
$UI.BTree.ShowLines = $true
Set-DoubleBuffered $UI.BTree
$UI.BSplit.Panel1.Controls.Add($UI.BTree)

$UI.DefBox = New-Object System.Windows.Forms.TextBox
$UI.DefBox.Dock = 'Fill'
$UI.DefBox.Multiline = $true
$UI.DefBox.ReadOnly = $true
$UI.DefBox.ScrollBars = 'Both'
$UI.DefBox.WordWrap = $false
$UI.DefBox.Font = New-Object System.Drawing.Font('Consolas', 10)
$UI.DefBox.Text = 'Use "Add instance..." to register a SQL Server, then expand it to browse databases, views and stored procedures.'
$UI.BBottom = New-Object System.Windows.Forms.Panel
$UI.BBottom.Dock = 'Bottom'
$UI.BBottom.Height = 42
$UI.BtnObjPlan = New-Control 'Button' 8 6 250 30 'Get plan for selected object...'
$UI.BtnCopyDef = New-Control 'Button' 266 6 150 30 'Copy definition'
$UI.BtnObjPlan.Enabled = $false
$UI.BtnCopyDef.Enabled = $false
$UI.BBottom.Controls.Add($UI.BtnObjPlan)
$UI.BBottom.Controls.Add($UI.BtnCopyDef)
$UI.BSplit.Panel2.Controls.Add($UI.DefBox)
$UI.BSplit.Panel2.Controls.Add($UI.BBottom)

$bstrip = New-Object System.Windows.Forms.ToolStrip
$bstrip.GripStyle = 'Hidden'
[void](Add-StripButton $bstrip 'Add instance...' { Invoke-AddInstance })
[void](Add-StripButton $bstrip 'Edit' { Invoke-EditInstance })
[void](Add-StripButton $bstrip 'Remove' { Invoke-RemoveInstance })
[void](Add-StripButton $bstrip 'Refresh' { Invoke-RefreshBrowser })
$UI.BrowserSys = Add-StripButton $bstrip 'Include system objects' { Reset-AllInstances } $true
$UI.TabBrowser.Controls.Add($UI.BSplit)
$UI.TabBrowser.Controls.Add($bstrip)

$UI.BTree.Add_BeforeExpand({
    param($sender, $e)
    if (Test-DummyNode $e.Node) {
        if (-not (Open-BrowserNode $e.Node)) { $e.Cancel = $true }
    }
})
$UI.BTree.Add_AfterSelect({
    param($sender, $e)
    Show-BrowserSelection $e.Node
})
$UI.BTree.Add_NodeMouseDoubleClick({
    param($sender, $e)
    $t = $e.Node.Tag
    if ($t -is [hashtable] -and $t.Kind -in @('View', 'Proc')) { Start-ObjectPlan }
})
$UI.BtnObjPlan.Add_Click({ Start-ObjectPlan })
$UI.BtnCopyDef.Add_Click({
    if ($UI.DefBox.Text) { try { [System.Windows.Forms.Clipboard]::SetText($UI.DefBox.Text) } catch { } }
})
$tabs.Add_SelectedIndexChanged({
    if ($tabs.SelectedTab -eq $UI.TabBrowser -and -not $script:BrowserSplitSet) {
        try { $UI.BSplit.SplitterDistance = 340; $script:BrowserSplitSet = $true } catch { }
    }
})

Import-Instances
foreach ($i in $script:Instances) { [void](Add-InstanceNode $i) }

# right: operator properties
$UI.Props = New-Object System.Windows.Forms.ListView
$UI.Props.Dock = 'Fill'
$UI.Props.View = 'Details'
$UI.Props.FullRowSelect = $true
$UI.Props.HideSelection = $false
$UI.Props.ShowGroups = $true
$UI.Props.HeaderStyle = 'None'
[void]$UI.Props.Columns.Add('Property', 170)
[void]$UI.Props.Columns.Add('Value', 300)
Set-DoubleBuffered $UI.Props
$UI.ValueBox = New-Object System.Windows.Forms.TextBox
$UI.ValueBox.Dock = 'Fill'
$UI.ValueBox.Multiline = $true
$UI.ValueBox.ReadOnly = $true
$UI.ValueBox.ScrollBars = 'Vertical'
$UI.ValueBox.Font = New-Object System.Drawing.Font('Consolas', 9.5)
$right.Panel1.Controls.Add($UI.Props)
$right.Panel2.Controls.Add($UI.ValueBox)

# ---------- events ----------

$canvas.Add_Paint({
    param($sender, $e)
    $st = $script:CurrentStmt
    if ($null -eq $st) {
        $e.Graphics.DrawString('Open a .sqlplan file, paste plan XML, drop a file here, or get a plan from SQL Server.', $UI.Font, [System.Drawing.Brushes]::Gray, [single]20, [single]20)
        return
    }
    $e.Graphics.TranslateTransform([single]$sender.AutoScrollPosition.X, [single]$sender.AutoScrollPosition.Y)
    $e.Graphics.ScaleTransform([single]$script:Zoom, [single]$script:Zoom)
    Invoke-PlanDraw $e.Graphics $st
})
$canvas.Add_Resize({ $UI.Canvas.Invalidate() })
$canvas.Add_Scroll({ $UI.Canvas.Invalidate() })
$canvas.Add_MouseEnter({ if ($UI.Form.ContainsFocus) { [void]$UI.Canvas.Focus() } })
$canvas.Add_MouseDown({
    param($sender, $e)
    [void]$UI.Canvas.Focus()
    $n = Get-NodeAt $e.X $e.Y
    if ($n) { Select-PlanNode $n }
})
$canvas.Add_MouseMove({
    param($sender, $e)
    $n = Get-NodeAt $e.X $e.Y
    if (-not [object]::ReferenceEquals($n, $script:HoverNode)) {
        $script:HoverNode = $n
        if ($n) {
            $UI.Canvas.Cursor = [System.Windows.Forms.Cursors]::Hand
            $UI.Tip.SetToolTip($UI.Canvas, (Get-NodeTip $n))
        } else {
            $UI.Canvas.Cursor = [System.Windows.Forms.Cursors]::Default
            $UI.Tip.SetToolTip($UI.Canvas, '')
        }
    }
})
$canvas.Add_MouseWheel({
    param($sender, $e)
    if (([System.Windows.Forms.Control]::ModifierKeys -band [System.Windows.Forms.Keys]::Control) -eq [System.Windows.Forms.Keys]::Control) {
        if ($e -is [System.Windows.Forms.HandledMouseEventArgs]) { $e.Handled = $true }
        if ($e.Delta -gt 0) { Set-Zoom ($script:Zoom + 0.1) } else { Set-Zoom ($script:Zoom - 0.1) }
    }
})

$UI.Stmts.Add_SelectedIndexChanged({
    if ($UI.Stmts.SelectedItems.Count -gt 0) { Select-Statement $UI.Stmts.SelectedItems[0].Tag }
})
$UI.Tree.Add_AfterSelect({
    param($sender, $e)
    if (-not $script:Syncing -and $e.Node -and $e.Node.Tag) { Select-PlanNode $e.Node.Tag -Scroll }
})
$UI.Top.Add_SelectedIndexChanged({
    if ($UI.Top.SelectedItems.Count -gt 0) { Select-PlanNode $UI.Top.SelectedItems[0].Tag -Scroll }
})
$UI.Find.Add_SelectedIndexChanged({
    if ($UI.Find.SelectedItems.Count -gt 0) {
        $fnd = $UI.Find.SelectedItems[0].Tag
        if ($fnd.Node) { Select-PlanNode $fnd.Node -Scroll }
    }
})
$UI.Miss.Add_SelectedIndexChanged({
    if ($UI.Miss.SelectedItems.Count -gt 0) { $UI.MissScript.Text = $UI.Miss.SelectedItems[0].Tag.Script }
})
$UI.Props.Add_SelectedIndexChanged({
    if ($UI.Props.SelectedItems.Count -gt 0) { $UI.ValueBox.Text = $UI.Props.SelectedItems[0].SubItems[1].Text }
})
$UI.Props.Add_Resize({
    $UI.Props.Columns[1].Width = [math]::Max(80, $UI.Props.ClientSize.Width - $UI.Props.Columns[0].Width - 4)
})

$form.Add_DragEnter({
    param($sender, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy }
})
$form.Add_DragDrop({
    param($sender, $e)
    $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    if ($files -and $files.Count -gt 0) { Open-PlanFile $files[0] }
})
$form.Add_KeyDown({
    param($sender, $e)
    if (-not $e.Control) { return }
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::O) { $e.SuppressKeyPress = $true; Invoke-OpenFile }
    elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::V) { $e.SuppressKeyPress = $true; Invoke-PasteXml }
    elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::R) { $e.SuppressKeyPress = $true; Start-QueryPlan }
    elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::Oemplus -or $e.KeyCode -eq [System.Windows.Forms.Keys]::Add) { $e.SuppressKeyPress = $true; Set-Zoom ($script:Zoom + 0.1) }
    elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::OemMinus -or $e.KeyCode -eq [System.Windows.Forms.Keys]::Subtract) { $e.SuppressKeyPress = $true; Set-Zoom ($script:Zoom - 0.1) }
})
[System.Windows.Forms.Application]::add_ThreadException({
    param($sender, $e)
    [void][System.Windows.Forms.MessageBox]::Show($e.Exception.Message, 'SQL Plan Viewer', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
})

$form.Add_FormClosed({ $script:Pw.Clear(); $script:LastQuery = '' })
$form.Add_Shown({
    try {
        $main.SplitterDistance = [int]($form.ClientSize.Width * 0.70)
        $left.SplitterDistance = 130
        $right.SplitterDistance = [int]($right.Height * 0.72)
    } catch { }
    $form.Activate()
    if ($PlanFile) { Open-PlanFile $PlanFile }
})

[void]$form.ShowDialog()

#endregion
