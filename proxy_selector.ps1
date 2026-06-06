# proxy_selector.ps1
# PowerShell script for Windows PowerShell 5.1
# Reads proxies from proxies.txt file

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$proxiesFile = Join-Path $scriptDir "proxies.txt"

# Check if proxies.txt exists
if (-not (Test-Path $proxiesFile)) {
    Clear-Host
    Write-Host "==================================================" -ForegroundColor Red
    Write-Host "         ERROR: proxies.txt not found!" -ForegroundColor Red
    Write-Host "==================================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please create proxies.txt file in the same folder" -ForegroundColor Yellow
    Write-Host "Format each line: IP:PORT|COUNTRY_CODE|COUNTRY_NAME" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Example:" -ForegroundColor Cyan
    Write-Host "  45.195.200.93:8080|WW|Worldwide" -ForegroundColor White
    Write-Host "  185.106.183.85:8080|SE|Sweden" -ForegroundColor White
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# Read proxies from file
$proxies = Get-Content $proxiesFile | Where-Object { $_.Trim() -ne "" -and $_ -notmatch "^#" }

if ($proxies.Count -eq 0) {
    Clear-Host
    Write-Host "==================================================" -ForegroundColor Red
    Write-Host "         ERROR: No proxies found in proxies.txt!" -ForegroundColor Red
    Write-Host "==================================================" -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

$countryMap = @{}
$countryProxies = @{}
$proxyToCountry = @{}

foreach ($entry in $proxies) {
    $parts = $entry -split '\|'
    if ($parts.Count -ne 3) {
        Write-Host "WARNING: Invalid line format: $entry" -ForegroundColor Yellow
        continue
    }
    
    $proxy = $parts[0]
    $code = $parts[1]
    $name = $parts[2]
    
    if (-not $countryProxies.ContainsKey($code)) {
        $countryProxies[$code] = @()
    }
    $countryProxies[$code] += $proxy
    $countryMap[$code] = $name
    $proxyToCountry[$proxy] = $name
}

if ($countryMap.Count -eq 0) {
    Clear-Host
    Write-Host "==================================================" -ForegroundColor Red
    Write-Host "         ERROR: No valid proxy entries!" -ForegroundColor Red
    Write-Host "==================================================" -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# Sort countries, but put WW (Worldwide) first
$sortedCodes = $countryMap.Keys | Where-Object { $_ -ne "WW" } | Sort-Object
$sortedCodes = @("WW") + $sortedCodes

# Parallel check using RunspacePool
function Test-ProxiesParallel {
    param($proxyList)
    
    $results = @()
    $runspacePool = [runspacefactory]::CreateRunspacePool(1, 20)
    $runspacePool.Open()
    
    $jobs = @()
    
    foreach ($proxy in $proxyList) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $runspacePool
        [void]$ps.AddScript({
            param($p)
            
            $testUrls = @(
                "http://www.google.com",
                "http://www.cloudflare.com",
                "http://www.microsoft.com"
            )
            
            foreach ($url in $testUrls) {
                try {
                    $uri = $url
                    $webRequest = [System.Net.WebRequest]::Create($uri)
                    $webRequest.Proxy = New-Object System.Net.WebProxy($p, $true)
                    $webRequest.Timeout = 4000
                    $webRequest.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
                    $webRequest.ServerCertificateValidationCallback = {$true}
                    
                    $startTime = Get-Date
                    $response = $webRequest.GetResponse()
                    $endTime = Get-Date
                    $response.Close()
                    
                    $latency = [math]::Round(($endTime - $startTime).TotalMilliseconds)
                    return @{Proxy = $p; Latency = $latency; Status = "OK"}
                }
                catch {
                    continue
                }
            }
            
            return @{Proxy = $p; Latency = "n/a"; Status = "FAIL"}
        }).AddArgument($proxy)
        
        $asyncResult = $ps.BeginInvoke()
        $jobs += @{
            PowerShell = $ps
            AsyncResult = $asyncResult
        }
    }
    
    foreach ($job in $jobs) {
        $result = $job.PowerShell.EndInvoke($job.AsyncResult)
        $results += $result
        $job.PowerShell.Dispose()
    }
    
    $runspacePool.Close()
    $runspacePool.Dispose()
    
    return $results
}

while ($true) {
    Clear-Host
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "         Proxy Country Selector" -ForegroundColor Green
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "[0] ALL PROXIES (all countries together)" -ForegroundColor Magenta
    Write-Host ""
    
    $i = 1
    $menuMap = @{}
    
    foreach ($code in $sortedCodes) {
        $countryName = $countryMap[$code]
        $count = $countryProxies[$code].Count
        Write-Host "[$i] $countryName ($count proxies)" -ForegroundColor Yellow
        $menuMap[$i] = $code
        $i++
    }
    
    Write-Host ""
    Write-Host "[q] Quit" -ForegroundColor Red
    Write-Host ""
    $choice = Read-Host "Select country (0-$($i-1))"
    
    if ($choice -eq 'q') {
        Write-Host "Goodbye!" -ForegroundColor Green
        Read-Host "Press Enter to exit"
        exit 0
    }
    
    # Handle ALL PROXIES
    if ($choice -eq '0') {
        Clear-Host
        Write-Host "==================================================" -ForegroundColor Cyan
        Write-Host "         ALL PROXIES (every country)" -ForegroundColor Green
        Write-Host "==================================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Checking all proxies (max 4 seconds each)..." -ForegroundColor Yellow
        Write-Host ""
        
        $allProxies = @()
        foreach ($code in $sortedCodes) {
            $allProxies += $countryProxies[$code]
        }
        
        if ($allProxies.Count -eq 0) {
            Write-Host "No available proxies." -ForegroundColor Red
            Read-Host "Press Enter to go back"
            continue
        }
        
        $startTotal = Get-Date
        $results = Test-ProxiesParallel -proxyList $allProxies
        $endTotal = Get-Date
        $totalTime = [math]::Round(($endTotal - $startTotal).TotalMilliseconds)
        
        $workingCount = ($results | Where-Object { $_.Status -eq "OK" }).Count
        $totalCount = $results.Count
        
        foreach ($result in $results) {
            $proxy = $result.Proxy
            $latency = $result.Latency
            $country = $proxyToCountry[$proxy]
            
            $output = "{0,-25} | {1} ... " -f $country, $proxy
            Write-Host -NoNewline $output
            if ($latency -eq "n/a") {
                Write-Host "n/a" -ForegroundColor Red
            } else {
                Write-Host "${latency}ms" -ForegroundColor Green
            }
        }
        
        Write-Host ""
        Write-Host "Working proxies: $workingCount / $totalCount" -ForegroundColor Cyan
        Write-Host "All checks completed in ${totalTime}ms" -ForegroundColor Cyan
        Write-Host ""
        Read-Host "Press Enter to return to menu"
        continue
    }
    
    # Handle regular country selection
    if (-not $menuMap.ContainsKey([int]$choice)) {
        Write-Host "Invalid choice!" -ForegroundColor Red
        Read-Host "Press Enter"
        continue
    }
    
    $selectedCode = $menuMap[[int]$choice]
    $selectedName = $countryMap[$selectedCode]
    $proxiesList = $countryProxies[$selectedCode] | Sort-Object { [version]($_ -replace ':\d+$','') }
    
    Clear-Host
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "         Proxies for $selectedName" -ForegroundColor Green
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Checking proxies (max 4 seconds each)..." -ForegroundColor Yellow
    Write-Host ""
    
    if ($proxiesList.Count -eq 0) {
        Write-Host "No available proxies." -ForegroundColor Red
        Read-Host "Press Enter to go back"
        continue
    }
    
    $startTotal = Get-Date
    $results = Test-ProxiesParallel -proxyList $proxiesList
    $endTotal = Get-Date
    $totalTime = [math]::Round(($endTotal - $startTotal).TotalMilliseconds)
    
    $workingCount = ($results | Where-Object { $_.Status -eq "OK" }).Count
    
    foreach ($result in $results) {
        $proxy = $result.Proxy
        $latency = $result.Latency
        
        Write-Host -NoNewline "$proxy ... "
        if ($latency -eq "n/a") {
            Write-Host "n/a" -ForegroundColor Red
        } else {
            Write-Host "${latency}ms" -ForegroundColor Green
        }
    }
    
    Write-Host ""
    Write-Host "Working proxies: $workingCount / $($proxiesList.Count)" -ForegroundColor Cyan
    Write-Host "All checks completed in ${totalTime}ms" -ForegroundColor Cyan
    Write-Host ""
    Read-Host "Press Enter to return to menu"
}