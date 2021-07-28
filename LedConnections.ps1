$host.UI.RawUI.WindowTitle = "LedConnections"
$OutputEncoding = [System.Console]::OutputEncoding = [System.Console]::InputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['*:Encoding'] = 'utf8'

$homeDir = "C:\SN_Scripts\LedConnections"
$jsonPath = "$homeDir\sn_data.json"
$startDate = get-date -DisplayHint date -Format dd/MM/yyyy
$mtd = "C:\Metadane_do_skryptow"

$excludedList = @()

$csv = @(
    "$mtd\meta_premium.csv",
    "$mtd\meta_city.csv",
    "$mtd\meta_super.csv",
    "$mtd\meta_pakiet.csv"
)

Function SendSlackMessage {
    param (
        [string] $message
    )

    $token = "***"
    $send = (Send-SlackMessage -Token $token -Channel 'led_connections' -Text $message).ok
    #$send = (Send-SlackMessage -Token $token -Channel 'testowanko' -Text $message).ok
    ( -join ("Wiadomosc wyslana: ", $send))
}

function GetComputersFromAPI {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipeline)]
        [ValidateNotNull()]
        [String]$networkName,
        [Array]$dontCheck
    )
    
    # API
    $requestURL = '***'
    $requestHeaders = @{'sntoken' = '***'; 'Content-Type' = 'application/json' }
    $requestBody = @"
{

"network": [$($networkName)]

}
"@

    # Request
    try {
        $request = Invoke-WebRequest -Uri $requestURL -Method POST -Body $requestBody -Headers $requestHeaders -ea Stop
    }
    catch [exception] {
        $Error[0]
        Exit 1
    }

    # Creating PS array of sn
    if ($request.StatusCode -eq 200) {
        $requestContent = $request.content | ConvertFrom-Json
    }
    else {
        Write-host ( -join ("Received bad StatusCode for request: ", $request.StatusCode, " - ", $request.StatusDescription)) -ForegroundColor Red
        Exit 1
    }

    $snList = @()
    $requestContent | ForEach-Object {
        if ((!($dontCheck -match $_.name)) -and ($_.lok -ne "LOK0014")) {
            $hash = [ordered]@{
                sn              = $_.name;
                ip              = $_.ip;
                lok_id          = $_.lok;
                placowka        = $_.lok_name.toString();
                sim             = "NULL";
            }

            $snList = [array]$snList + (New-Object psobject -Property $hash)
        }
    }

    return $snList
}

function Start-SleepTimer($seconds) {
    $doneDT = (Get-Date).AddSeconds($seconds)
    while ($doneDT -gt (Get-Date)) {
        $secondsLeft = $doneDT.Subtract((Get-Date)).TotalSeconds
        $percent = ($seconds - $secondsLeft) / $seconds * 100
        Write-Progress -activity "LED connections" -Status "Nastepne sprawdzenie polaczen za" -SecondsRemaining $secondsLeft -PercentComplete $percent
        [System.Threading.Thread]::Sleep(500)
    } 

    Write-Progress -activity "Start-sleep" -Status "Nastepne sprawdzenie polaczen za" -SecondsRemaining 0 -Completed
}

function Psql {
    param (
        $request
    )

    $psqlServer = "localhost"
    $psqlPort = 5432
    $psqlDB = "***"
    $psqlUid = "***"
    $psqlPass = "***"

    try {
        $DBConnectionString = "Driver={PostgreSQL UNICODE(x64)};Server=$psqlServer;Port=$psqlPort;Database=$psqlDB;Uid=$psqlUid;Pwd=$psqlPass;ConnSettings=SET CLIENT_ENCODING TO 'WIN1250';"
        $DBConn = New-Object System.Data.Odbc.OdbcConnection;
        $DBConn.ConnectionString = $DBConnectionString;
        $DBConn.Open();
        $DBCmd = $DBConn.CreateCommand();
        $DBCmd.CommandText = $request;
        $rdr = $DBCmd.ExecuteReader()
        $tbl = New-Object Data.DataTable
        $tbl.Load($rdr)
        $rdr.Close()
    }
    catch {
        Write-host "Error while connecting to database: $($_.exception.message)" -ForegroundColor Red
        $tbl = "error"
    }

    return $tbl
}

Get-Content "$mtd\sn_disabled_LedConnections.txt" -ErrorAction SilentlyContinue | ForEach-Object { 
    if ($excludedList -notcontains $_) {
        $excludedList = [array]$excludedList + $_
    }
}

while (!(Test-Connection -ComputerName "10.99.99.10" -Count 3 -Quiet)) {
    Write-Host "VPN not connected!" -ForegroundColor Red
    Start-Sleep -Seconds 30
}

do {
    $dateNow = get-date -DisplayHint date -Format dd/MM/yyyy
    $freshData = GetComputersFromAPI -networkName '"LED City", "LED Premium", "Super Screen", "Pakiet Tranzyt"' -dontCheck $excludedList
    $csvContent = @()

    foreach ($c in $csv) {
        $csvContent = $csvContent + (Import-Csv -Path $c -Delimiter ';' | Select-Object 'ID', 'Restart_SMS')
    }

    if (Test-Path $jsonPath) { 
        try { 
            [System.Collections.ArrayList]$localData = ConvertFrom-Json (Get-Content $jsonPath -Raw -ea Continue) -ea Continue 
            $runUpdate = $true
        }
        catch {
            Write-Host "ERROR: $($_.Exception.message)" 
        }
    }
    else {
        ConvertTo-Json -InputObject $freshData | Out-File $jsonPath
        $runUpdate = $false
    }

    # JSON UPDATE
    if ($runUpdate) {
        foreach ($f in $freshData) {
            $counter = 0
            $ldCount = $localData.Count
            For ($i = 0; $i -lt $ldCount; $i++) {
                if ($f.sn -eq $localData[$i].sn) {
                    # IP update
                    if (($f.ip -ne $localData[$i].ip) -and ($f.ip -ne "NULL") -and ($f.ip -ne "")) {
                        $localData[$i].ip = $f.ip
                    }

                    # LOK_ID update
                    if (($f.lok_id -ne $localData[$i].lok_id) -and ($f.lok_id -ne "null") -and ($f.lok_id -ne "")) {
                        $localData[$i].lok_id = $f.lok_id
                    }

                    # PLACOWKA update
                    if (($f.placowka -ne $localData[$i].placowka) -and ($f.placowka -ne "null") -and ($f.placowka -ne "")) {
                        $localData[$i].placowka = $f.placowka
                    }
                }
                else {
                    $counter++
                    if ($counter -eq $ldCount) {
                        # ADD NEW SN
                        $hash = [ordered]@{
                            sn              = $f.sn;
                            ip              = $f.ip;
                            lok_id          = $f.lok_id;
                            placowka        = $f.placowka;
                            sim             = "NULL";
                        }
                
                        Write-host "Adding $($f.sn) to json"
                        $localData = [array]$localData + (New-Object psobject -Property $hash)
                    }
                }
            }
        }

        for ($l = 0; $l -lt $localData.count; $l++) {
            $n = $localData[$l].sn
        
            # REMOVE MISSING SN
            if (!($freshData.sn -contains $n)) {
                Write-host "Removing $n from json"
                $localData.Remove($localData[$l])
            }

            # SIM NUMBER UPDATE
            if ($csvContent.ID -contains $localData[$l].lok_id) {
                $simNumber = ($csvContent | Where-Object { $_.ID -eq $localData[$l].lok_id }).Restart_SMS
            
                if (($simNumber -eq "N/A") -or ($simNumber -eq "") -or ($null -eq $simNumber)) {
                    $localData[$l].sim = "null"
                }
                else {
                    $localData[$l].sim = $simNumber
                }
            }
        }

        ConvertTo-Json -InputObject $localData | Out-File $jsonPath -Force
    }

    [System.Collections.ArrayList]$servers = ConvertFrom-Json (Get-Content $jsonPath -Raw -ea Continue) -ea Continue
    
    for ($i = 0; $i -lt $servers.Count; $i++) {
        $scriptBlock = {
            $OutputEncoding = [System.Console]::OutputEncoding = [System.Console]::InputEncoding = [System.Text.Encoding]::UTF8
            $PSDefaultParameterValues['*:Encoding'] = 'utf8bom'
            $serv = $args[0]
            $l = $args[1]
            $sn = $serv[$l].sn
            $snIP = $serv[$l].ip
            $lok = $serv[$l].placowka
            $lok_id = $serv[$l].lok_id

            Function SendSlackMessage {
                param (
                    [string] $message
                )
            
                $token = "***"
                $send = (Send-SlackMessage -Token $token -Channel 'led_connections' -Text $message).ok
                #$send = (Send-SlackMessage -Token $token -Channel 'testowanko' -Text $message).ok
                ( -join ("Wiadomosc wyslana: ", $send))
            }

            function Psql {
                param (
                    $request
                )
            
                $psqlServer = "localhost"
                $psqlPort = 5432
                $psqlDB = "***"
                $psqlUid = "***"
                $psqlPass = "***"
            
                try {
                    $DBConnectionString = "Driver={PostgreSQL UNICODE(x64)};Server=$psqlServer;Port=$psqlPort;Database=$psqlDB;Uid=$psqlUid;Pwd=$psqlPass;ConnSettings=SET CLIENT_ENCODING TO 'WIN1250';"
                    $DBConn = New-Object System.Data.Odbc.OdbcConnection;
                    $DBConn.ConnectionString = $DBConnectionString;
                    $DBConn.Open();
                    $DBCmd = $DBConn.CreateCommand();
                    $DBCmd.CommandText = $request;
                    $rdr = $DBCmd.ExecuteReader()
                    $tbl = New-Object Data.DataTable
                    $tbl.Load($rdr)
                    $rdr.Close()
                }
                catch {
                    Write-host "Error while connecting to database: $($_.exception.message)" -ForegroundColor Red
                    $tbl = "error"
                }
            
                return $tbl
            }

            $db = Psql -request "SELECT * FROM led_connections;"
            
            if ((!($db.lok_id -contains $lok_id)) -and $db -ne "error") {
                Psql -request "INSERT INTO led_connections (led_name, lok_id) VALUES ('$lok', '$lok_id');"
            }

            if ($excludedList -notcontains $sn) {
                Write-Output "Checking connection with: $sn, $lok"
                $q = Psql -request "SELECT last_chk FROM led_connections WHERE lok_id = '$lok_id';"

                if (Test-Connection -ComputerName $snIP -Count 3 -Quiet) {
                    #maszyna pinguje
                    Write-Host "Connected" -ForegroundColor Green

                    if ($q.last_chk -eq 1) {
                        SendSlackMessage -message "*$sn*, $lok - jest znowu polaczony"
                        Psql -request "UPDATE led_connections SET last_chk = 0 WHERE lok_id = '$lok_id';"
                    }
                }
                else {
                    #maszyna nie pinguje
                    Write-Host "Not connected" -ForegroundColor Red  
                    Start-Sleep -Seconds 540

                    #ponowne sprawdzenie polaczenia
                    Write-Output "Re-checking connection with: $sn"
                    $q = Psql -request "SELECT disc_cntr, last_chk FROM led_connections WHERE lok_id = '$lok_id';"
        
                    if (Test-Connection -ComputerName $snIP -Count 3 -Quiet) {
                        #maszyna pinguje
                        Write-Host "Connected" -ForegroundColor Green

                        if ($q.last_chk -eq 1) {
                            SendSlackMessage -message "*$sn*, $lok - jest znowu polaczony"
                            Psql -request "UPDATE led_connections SET last_chk = 0 WHERE lok_id = '$lok_id';"
                        }
                    }
                    else {
                        #maszyna nie pinguje
                        Write-Host "Not connected" -ForegroundColor Red  
                        
                        if (($q.last_chk -eq 0) -and ($q -ne "error")) {
                            $msg = "*``$sn, $lok- jest niepolaczony!``*"
                            $simNumber = $serv[$l].sim

                            if (($null -eq $simNumber) -or ($simNumber -eq "null")) {
                                $simNumber = "Brak numeru SIM w pliku CSV"
                            }

                            $msg = "$msg`n sim: $simNumber" 
                            SendSlackMessage -message $msg                
                            $newCntr = $q.disc_cntr + 1
                            Psql -request "UPDATE led_connections SET disc_cntr = $newCntr, last_disc = NOW(), last_chk = 1 WHERE lok_id = '$lok_id';"
                        }
                    }
                }

                ""
            }
        }

        Start-Job -ScriptBlock $scriptBlock -Name $sn -ArgumentList $servers, $i
    }

    # Wait all jobs
    Get-Job | Wait-Job

    # Receive all jobs
    Get-Job | Receive-Job

    # Remove all jobs
    Get-Job | Remove-Job

    ConvertTo-Json -InputObject $servers | Out-File $jsonPath -Force
    Start-SleepTimer 180
}  until ($dateNow -gt $startDate)

$finalResult = @()
$db = Psql -request "SELECT * FROM led_connections;"
foreach ($r in $db) { 
    if (($r.last_chk -eq 1) -and ($servers.lok_id -contains $r.lok_id)) {
        $servers | ? { $_.lok_id -eq $r.lok_id } | % { $finalResult += $_ }
    }
}

if ($finalResult.count -ne 0) {
    $msg = "*Niepolaczone komputery na koniec dnia - $($startDate)* ``````` - SN - LOKALIZACJA -`n**********************`n"

    foreach ($result in $finalResult) {
        $msg += ( -join ($result.sn, " - ", $result.placowka, "`n"))
    }

    $msg += "``````` "
    SendSlackMessage -message $msg
}
else {
    SendSlackMessage -message "*Brak niepolaczonych komputerow - $($startDate)*"
}