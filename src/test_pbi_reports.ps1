<#
    Test Runner for Power BI crash tests
    This script will try to open each specified Power BI report (default is *.pbit)
    and detect whether Power BI can load it. If any error messages are detected or
    the loading times out, a failure is detected and reported. Additionally, all
    opened Power BI windows are screenshotted and stored in the output directory.
    Copyright © Christoph Thiede 2020.
#>
[CmdletBinding()]
Param(
    [ValidateScript({Test-Path $_})]
    $Reports = (Get-ChildItem *.pbit),
    [ValidateScript({Test-Path $_})]
    [string] $OutputPath = (Join-Path $PSScriptRoot "output"),

    [Parameter(HelpMessage="PBI Desktop executable file.")]
    $Pbi = (Join-Path ${env:LOCALAPPDATA} "Microsoft\WindowsApps\PBIDesktopStore.exe"),

    [Parameter(HelpMessage="If loading time is exceeded, report test will abort and fail.")]
    [timespan] $Timeout = [timespan]::FromSeconds(120),
    [Parameter(HelpMessage="Delay between checks.")]
    [timespan] $CheckInterval = [timespan]::FromSeconds(10),
    [Parameter(HelpMessage="Time to wait for PBI to show the initial splash screen.")]
    [timespan] $PreLoadDelay = [timespan]::FromSeconds(15),
    [Parameter(HelpMessage="Time to wait for PBI to show possible loading windows after model " +
        "has been created.")]
    [timespan] $LoadDelay = [timespan]::FromSeconds(15)
)


if ($env:CI) {
    # Forward Write-Progress messages to stdout on CI
    . "src/utils/Write-Progress-Stdout.ps1"
}

$csharpProvider = (&"src/utils/csharp_provider.ps1")[-1]
$assemblies = (
    "nuget\NETStandard.Library.2.0.0\build\netstandard2.0\ref\mscorlib.dll",
    "nuget\NETStandard.Library.2.0.0\build\netstandard2.0\ref\netstandard.dll",
    "nuget\NETStandard.Library.2.0.0\build\netstandard2.0\ref\System.Drawing.dll",
    "nuget\Magick.NET-Q8-AnyCPU.7.19.0\lib\netstandard20\Magick.NET-Q8-AnyCPU.dll",
    "nuget\Magick.NET.Core.2.0.0\lib\netstandard20\Magick.NET.Core.dll"
) | ForEach-Object {[System.Reflection.Assembly]::LoadFrom($_)}
Copy-Item "nuget\Magick.NET-Q8-AnyCPU.7.19.0\runtimes\win-x86\native\Magick.Native-Q8-x86.dll" .
Copy-Item "nuget\Magick.NET-Q8-AnyCPU.7.19.0\runtimes\win-x64\native\Magick.Native-Q8-x64.dll" .
Add-Type `
    -CodeDomProvider $csharpProvider `
    -ReferencedAssemblies $assemblies `
    -TypeDefinition (Get-Content -Path src/test_pbi_reports.cs | Out-String)
if (!$?) {
    exit 2  # Error loading C# component
}
$testClass = [PbiCrashTests.PbiReportTestCase]


$maxIntervalCount = [Math]::Ceiling($Timeout.TotalMilliseconds / $CheckInterval.TotalMilliseconds)
$Timeout = [timespan]::FromMilliseconds($CheckInterval.TotalMilliseconds * $maxIntervalCount)

$global:runs = New-Object System.Collections.Generic.List[$testClass]
$global:passes = New-Object System.Collections.Generic.List[$testClass]
$global:failures = New-Object System.Collections.Generic.List[$testClass]
$global:errors = New-Object System.Collections.Generic.List[$testClass]

function Invoke-Test([PbiCrashTests.PbiReportTestCase] $test) {
    Write-Progress -Id 2 -Activity "Testing report" -CurrentOperation "Opening report file"
    $test.Start()

    try {
        $startTime = Get-Date
        for ($i = 1;; $i++) {
            $elapsed = (Get-Date) - ($startTime)
            if ($elapsed -gt $Timeout) {
                Write-Error "⚠ TIMEOUT: $test"
                $global:errors.Add($test)
                return
            }

            Write-Progress -Id 2 -Activity "Testing report" `
                -CurrentOperation "Waiting for report file to load... ($i/$maxIntervalCount)"
            Start-Sleep -Seconds $CheckInterval.Seconds
            $test.Check()

            if ($test.HasPassed) {
                Write-Output "✅ PASS: $test"
                $global:passes.Add($test)
                return
            } elseif ($test.HasFailed) {
                $err = @("❌ FAILED: $test")
                if ($test.ResultReason) {
                    $err += $test.ResultReason
                }
                Write-Error ($err -join "`n")
                $global:failures.Add($test)
                return
            }
        }
    } finally {
        $test.SaveResults($OutputPath)
        Write-Progress -Id 2 -Activity "Testing report" -CurrentOperation "Closing report file"
        $test.Stop()
        Write-Progress -Id 2 -Completed "Testing report"
    }
}


# Prepare test cases
$tests = $Reports | ForEach-Object {[PbiCrashTests.PbiReportTestCase]::new(
        $_, $Pbi, $PreLoadDelay, $LoadDelay
    )}
mkdir -Force $OutputPath | Out-Null

# Run tests
foreach ($test in $tests) {
    $runs.Add($test)
    Write-Progress -Id 1 -Activity "Testing Power BI reports" `
        -CurrentOperation $test -PercentComplete ($runs.Count / $tests.Count)
    Invoke-Test $test
}
Write-Progress -Id 1 -Completed "Testing Power BI reports"

# Summary
Write-Output "`nPower BI Test summary:"
if ($runs) {
    @('passes', 'failures', 'errors') | ForEach-Object {
        $testGroup = (Get-Variable $_).Value
        $runs = [Linq.Enumerable]::ToList([Linq.Enumerable]::Except($runs, $testGroup))
        [PSCustomObject]@{
            Group = $_
            Count = $testGroup.Count
            Tests = (($testGroup | ForEach-Object Name) -join ', ')
    }} | Where-Object {$_.Count}
} else {
    Write-Output "No tests have been executed."
}
if ($runs) {
    Write-Error "Warning: $($runs.Count) tests have an unknown result: $runs"
}
$unsuccessful = $failures + $errors
Write-Output "Screenshots of all opened reports have been stored in $OutputPath."

exit !!($unsuccessful)
