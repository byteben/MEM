$results = foreach ($context in Get-ChildItem "HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps" -ErrorAction SilentlyContinue) {

    $grsPath = Join-Path $context.PSPath "GRS"
    if (-not (Test-Path $grsPath)) { continue }

    foreach ($subgraph in Get-ChildItem $grsPath -ErrorAction SilentlyContinue) {

        $vals = $subgraph.GetValueNames()

        if ($vals.Count -eq 1 -and $vals[0] -eq "SubgraphEvaluationTimeUTC") {

            try {
                $time = (Get-ItemProperty -LiteralPath $subgraph.PSPath -Name SubgraphEvaluationTimeUTC -ErrorAction Stop).SubgraphEvaluationTimeUTC

                [pscustomobject]@{
                    Context  = $context.PSChildName
                    Subgraph = $subgraph.PSChildName
                    EvalTime = $time
                    AgeDays  = [int]((Get-Date).ToUniversalTime() - ([datetime]$time)).TotalDays
                }
            } catch {}
        }
    }
}

"Total orphaned GRS entries: $($results.Count)"
$results | Sort-Object AgeDays -Descending | Format-Table -AutoSize