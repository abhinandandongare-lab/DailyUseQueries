<#
.SYNOPSIS
    Finds the call chain(s) between a parent stored procedure and a child
    stored procedure by statically analysing the SQL source files.

.DESCRIPTION
    This tool assumes each stored procedure lives in its own source file whose
    name (without extension) is the procedure name, somewhere under a folder
    tree. Procedures call each other with EXEC / EXECUTE <proc>. A child
    procedure is often not called directly by the parent, but reached through
    one or more intermediate procedures.

    This script:
      1. Scans every matching file under -Path and extracts the procedure name
         from the file name.
      2. Strips comments and extracts every EXEC/EXECUTE target.
      3. Builds a call graph (parent proc -> procedures it calls).
      4. Searches the graph for path(s) that start at -ParentProc and reach
         -ChildProc, printing each hop in the chain.

    The graph is cached to disk so repeated queries are fast. Use -Refresh to
    rebuild it after the source changes.

.PARAMETER ParentProc
    Name of the starting (parent) procedure. The procedure-name suffix (e.g.
    _sp) and any schema prefix are optional.

.PARAMETER ChildProc
    Name of the target (child) procedure to reach.

.PARAMETER Path
    Root folder to scan recursively for procedure source files. (Alias: -Root)

.PARAMETER Extension
    File pattern of the procedure source files to scan. Default '*.sp'.
    Use '*.sql' for projects that keep each procedure in a .sql file.

.PARAMETER AllPaths
    Return every distinct path (up to -MaxDepth), not just the shortest one.

.PARAMETER MaxDepth
    Maximum chain length to search when -AllPaths is used. Default 25.

.PARAMETER Refresh
    Rebuild the cached call graph instead of reusing it.

.EXAMPLE
    # Shortest chain between two procedures
    .\Find-SpCallChain.ps1 -ParentProc parent_proc_sp -ChildProc child_proc_sp -Path 'C:\src\MyDb'

.EXAMPLE
    # Positional arguments (parent, child) with an explicit folder
    .\Find-SpCallChain.ps1 parent_proc_sp child_proc_sp -Path .\sql

.EXAMPLE
    # List every possible chain, not just the shortest one
    .\Find-SpCallChain.ps1 -ParentProc parent_proc_sp -ChildProc child_proc_sp -Path .\sql -AllPaths

.EXAMPLE
    # Projects that store each procedure in a .sql file, limit search depth
    .\Find-SpCallChain.ps1 parent_proc child_proc -Path .\sql -Extension '*.sql' -AllPaths -MaxDepth 10

.EXAMPLE
    # Rebuild the cached call graph after the source files changed
    .\Find-SpCallChain.ps1 parent_proc_sp child_proc_sp -Path .\sql -Refresh
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$ParentProc,

    [Parameter(Mandatory = $true, Position = 1)]
    [string]$ChildProc,

    [Parameter(Mandatory = $true)]
    [Alias('Root')]
    [string]$Path,

    [string]$Extension = '*.sp',

    [switch]$AllPaths,

    [int]$MaxDepth = 25,

    [switch]$Refresh
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Get-ShortName {
    param([string]$Name)
    # Strip schema/owner prefix and surrounding brackets, lower case.
    $n = $Name.Trim()
    $n = $n -replace '[\[\]]', ''
    $n = ($n -split '\.')[-1]
    return $n.ToLowerInvariant()
}

function Remove-SqlComments {
    param([string]$Text)
    # Remove /* ... */ block comments (non greedy, across lines).
    $t = [regex]::Replace($Text, '/\*.*?\*/', ' ', 'Singleline')
    # Remove -- line comments.
    $t = [regex]::Replace($t, '--[^\r\n]*', ' ')
    return $t
}

# ---------------------------------------------------------------------------
# Build (or load) the call graph
# ---------------------------------------------------------------------------
$rootName  = [IO.Path]::GetFileName($Path.TrimEnd('\', '/'))
$extTag    = ($Extension -replace '[^A-Za-z0-9]', '')
$cachePath = Join-Path $PSScriptRoot ('SpCallGraph_{0}_{1}.json' -f $rootName, $extTag)

if (-not $Refresh -and (Test-Path $cachePath)) {
    Write-Host "Loading cached call graph from $cachePath" -ForegroundColor DarkGray
    $cache     = Get-Content -LiteralPath $cachePath -Raw | ConvertFrom-Json
    $graph     = @{}
    $fileOf    = @{}
    foreach ($p in $cache.graph.PSObject.Properties) { $graph[$p.Name]  = @($p.Value) }
    foreach ($p in $cache.files.PSObject.Properties) { $fileOf[$p.Name] = $p.Value }
}
else {
    Write-Host "Scanning $Extension files under $Path ..." -ForegroundColor Cyan
    $spFiles = Get-ChildItem -LiteralPath $Path -Recurse -Filter $Extension -File -ErrorAction SilentlyContinue

    Write-Host ("Found {0} procedure files. Building call graph..." -f $spFiles.Count) -ForegroundColor Cyan

    # procKey = short procedure name (lower case). Map to file path.
    $fileOf = @{}
    foreach ($f in $spFiles) {
        $key = Get-ShortName ([IO.Path]::GetFileNameWithoutExtension($f.Name))
        if (-not $fileOf.ContainsKey($key)) { $fileOf[$key] = $f.FullName }
    }

    $execRegex = [regex]'(?i)\bEXEC(?:UTE)?\s+(?:@\w+\s*=\s*)?(?:\[?\w+\]?\.)?(\[?[A-Za-z_][\w]*\]?)'

    $graph = @{}
    $i = 0
    foreach ($f in $spFiles) {
        $i++
        if ($i % 1000 -eq 0) { Write-Host "  processed $i / $($spFiles.Count)" -ForegroundColor DarkGray }

        $procKey = Get-ShortName ([IO.Path]::GetFileNameWithoutExtension($f.Name))
        if (-not $graph.ContainsKey($procKey)) { $graph[$procKey] = @{} }

        $raw  = Get-Content -LiteralPath $f.FullName -Raw -ErrorAction SilentlyContinue
        if (-not $raw) { continue }
        $code = Remove-SqlComments $raw

        foreach ($m in $execRegex.Matches($code)) {
            $target = Get-ShortName $m.Groups[1].Value
            # Only keep edges to procedures that actually exist in the code base
            # and ignore self references.
            if ($target -ne $procKey -and $fileOf.ContainsKey($target)) {
                $graph[$procKey][$target] = $true
            }
        }
    }

    # Flatten edge sets to arrays for storage / traversal.
    $flat = @{}
    foreach ($k in $graph.Keys) { $flat[$k] = @($graph[$k].Keys) }
    $graph = $flat

    # Persist cache.
    $cacheObj = [pscustomobject]@{
        root  = $Path
        built = (Get-Date).ToString('o')
        graph = $graph
        files = $fileOf
    }
    $cacheObj | ConvertTo-Json -Depth 6 -Compress | Set-Content -LiteralPath $cachePath -Encoding UTF8
    Write-Host "Call graph cached to $cachePath" -ForegroundColor DarkGray
}

# ---------------------------------------------------------------------------
# Resolve input names and search
# ---------------------------------------------------------------------------
$start = Get-ShortName $ParentProc
$goal  = Get-ShortName $ChildProc

if (-not $fileOf.ContainsKey($start)) {
    Write-Warning "Parent procedure '$ParentProc' was not found as a $Extension file under $Path."
    return
}
if (-not $fileOf.ContainsKey($goal)) {
    Write-Warning "Child procedure '$ChildProc' was not found as a $Extension file under $Path."
    return
}

function Get-Edges {
    param([string]$Node)
    # Force an array so a node with a single edge is not unwrapped to a scalar.
    if ($graph.ContainsKey($Node)) { return ,@($graph[$Node]) }
    return ,@()
}

function Format-Chain {
    param([string[]]$Path)
    $lines = for ($n = 0; $n -lt $Path.Count; $n++) {
        $indent = '  ' * $n
        $arrow  = if ($n -eq 0) { '' } else { '-> ' }
        $name   = $Path[$n]
        $file   = $fileOf[$name]
        '{0}{1}{2}    [{3}]' -f $indent, $arrow, $name, $file
    }
    return $lines -join "`n"
}

Write-Host ''
Write-Host ("Searching for chain: {0}  ==>  {1}" -f $start, $goal) -ForegroundColor Green
Write-Host ''

if ($AllPaths) {
    # Depth limited DFS enumerating all simple paths.
    $results = New-Object System.Collections.Generic.List[object]
    $stack   = New-Object System.Collections.Generic.Stack[object]
    $stack.Push([pscustomobject]@{ Node = $start; Path = @($start) })

    while ($stack.Count -gt 0) {
        $cur = $stack.Pop()
        if ($cur.Node -eq $goal) { $results.Add($cur.Path); continue }
        if ($cur.Path.Count -ge $MaxDepth) { continue }
        foreach ($next in Get-Edges $cur.Node) {
            if ($cur.Path -notcontains $next) {
                $stack.Push([pscustomobject]@{ Node = $next; Path = $cur.Path + $next })
            }
        }
    }

    if ($results.Count -eq 0) {
        Write-Warning "No call chain found from '$start' to '$goal' (within depth $MaxDepth)."
        return
    }

    $results = $results | Sort-Object { $_.Count }
    Write-Host ("Found {0} chain(s):" -f $results.Count) -ForegroundColor Green
    $idx = 0
    foreach ($p in $results) {
        $idx++
        Write-Host ''
        Write-Host ("--- Chain #{0} ({1} hops) ---" -f $idx, ($p.Count - 1)) -ForegroundColor Yellow
        Write-Host (Format-Chain $p)
    }
}
else {
    # BFS for the shortest chain.
    $queue   = New-Object System.Collections.Generic.Queue[object]
    $visited = @{ $start = $true }
    $queue.Enqueue(@($start))
    $found   = $null

    while ($queue.Count -gt 0 -and -not $found) {
        $curPath = $queue.Dequeue()
        $node = $curPath[-1]
        foreach ($next in Get-Edges $node) {
            if ($next -eq $goal) { $found = $curPath + $next; break }
            if (-not $visited.ContainsKey($next)) {
                $visited[$next] = $true
                $queue.Enqueue($curPath + $next)
            }
        }
    }

    if (-not $found) {
        Write-Warning "No call chain found from '$start' to '$goal'. (Try -AllPaths or check the names.)"
        return
    }

    Write-Host ("Shortest chain ({0} hops):" -f ($found.Count - 1)) -ForegroundColor Yellow
    Write-Host ''
    Write-Host (Format-Chain $found)
    Write-Host ''
    Write-Host "(Use -AllPaths to list every possible chain.)" -ForegroundColor DarkGray
}
