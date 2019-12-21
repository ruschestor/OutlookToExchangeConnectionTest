# Script: Outlook to Exchange Connection Test
# Version: 1 (21.12.2019)
# Blog: https://itgeeknotes.blogspot.com

Remove-Variable * -ErrorAction SilentlyContinue


# VARIABLES
$version = 1

$date = (get-date).tostring("yyyy-MM-dd_HH-mm-ss")
$date2 = "Outlook to Exchange Connection Test | Version $version | " + (get-date).tostring("yyyy.MM.dd HH:mm:ss")

$path = "C:/temp/OutlookToExchangeConnectionTest"
$HTMLReportFile = "$path/OtECT_$date.html"

$Exchange = "exchange.test.local" # <<<<<<<<<------------------------------------ FILL HERE FQDN OF YOUR EXCHANGE SERVER
$SMTPDomain = "test.local"        # <<<<<<<<<------------------------------------ FILL HERE YOUR EMAIL DOMAIN NAME

$Exchange_EMSMDB_URL = "https://$Exchange/mapi/emsmdb/?showdebug=yes"
$Exchange_MAPIHC_URL = "https://$Exchange/mapi/healthcheck.htm"
$Exchange_EWSHC_URL = "https://$Exchange/ews/healthcheck.htm"
$Exchange_ADHC_URL = "https://$Exchange/autodiscover/healthcheck.htm"
$Exchange_AD_URL1 = "https://autodiscover.$SMTPDomain/autodiscover/autodiscover.xml"
$Exchange_AD_URL2 = "https://$SMTPDomain/autodiscover/autodiscover.xml"

$hosts_file = "C:\Windows\System32\drivers\etc\hosts"
$errorcolor = " bgcolor='#F84A42'"

$OutlookPaths = @("C:\Program Files\Microsoft Office\Office14\outlook.exe","C:\Program Files\Microsoft Office\Office15\outlook.exe","C:\Program Files\Microsoft Office\Office16\outlook.exe","C:\Program Files (x86)\Microsoft Office\Office14\outlook.exe","C:\Program Files (x86)\Microsoft Office\Office15\outlook.exe","C:\Program Files (x86)\Microsoft Office\Office16\outlook.exe","C:\Program Files (x86)\Microsoft Office\root\Office14\outlook.exe","C:\Program Files (x86)\Microsoft Office\root\Office15\outlook.exe","C:\Program Files (x86)\Microsoft Office\root\Office16\outlook.exe","C:\Program Files (x86)\Microsoft Office 14\ClientX86\Root\Office14\outlook.exe","C:\Program Files (x86)\Microsoft Office 16\ClientX86\Root\Office16\outlook.exe","C:\Program Files (x86)\Microsoft Office 15\ClientX86\Root\Office15\outlook.exe","C:\Program Files\Microsoft Office 14\ClientX64\Root\Office14\outlook.exe","C:\Program Files\Microsoft Office 15\ClientX64\Root\Office15\outlook.exe","C:\Program Files\Microsoft Office 16\ClientX64\Root\Office16\outlook.exe")
$OutlookRegAD = @("HKCU:\Software\Microsoft\Office\14.0\Outlook\AutoDiscover","HKCU:\Software\Microsoft\Office\15.0\Outlook\AutoDiscover","HKCU:\Software\Microsoft\Office\16.0\Outlook\AutoDiscover")

$EventLog_Minutes = 10

# FUNCTIONS
function GetIP
{
	Param ([string]$ip_req)
	$ip_result = (Resolve-DnsName -Name $ip_req -Type A -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).IP4Address
	return $ip_result
}

function GetLatency
{
	Param ([string]$latency_req)
    $Intermediate_result = Test-Connection -IPAddress $latency_req -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -Count 2
    if ($Intermediate_result -ne $null) {
	    $latency_result = [math]::Round(($Intermediate_result | Measure-Object -Property ResponseTime -Average).Average)
    }
    else
    {
        $latency_result = "error"
    }
	return $latency_result
}

function GetPort
{
	Param ([string]$GetPort_ip, [string]$GetPort_port)
	$GetPort_result = (Test-NetConnection -Port $GetPort_port -ComputerName $GetPort_ip -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).TcpTestSucceeded
    $GetPort_result = ":" + $GetPort_port + " - " + $GetPort_result
	return $GetPort_result
}

function GetURL
{
	Param ([string]$GetURL_req)

    $HTTP_Request = [System.Net.WebRequest]::Create($GetURL_req)
    $HTTP_Response = $HTTP_Request.GetResponse()
    if ($? -eq $false) {
	    $HTTP_error_text = $error[0]
        if ($HTTP_error_text -match "(401)") {$GetURL_result = "ERROR 401 - Unauthorized"}
    }
    $HTTP_Status = [int]$HTTP_Response.StatusCode
    if ($HTTP_Status -eq 200) {
        $GetURL_result = (Invoke-WebRequest -Uri $GetURL_req -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).RawContent
    }
    $HTTP_Response.Close()
	return $GetURL_result
}

function GetHTML ([string]$GetHTML_req)
{
	$GetHTML_result = (((($GetHTML_req.replace("`r`n","<br>")).replace("<br><br><br><br>","<br>")).replace("<br><br><br>","<br>")).replace("<br><br>","<br>"))
    if ($GetHTML_result.StartsWith("<br>")) {$GetHTML_result = $GetHTML_result -replace "^<br>", ""}
    if ($GetHTML_result -match "<br>$") {$GetHTML_result = $GetHTML_result -replace "<br>$", ""}
	return $GetHTML_result
}

###############################################################################################

# Create folder for report
If(!(test-path $path))
{
      New-Item -ItemType Directory -Force -Path $path
}
$IP = ((ipconfig | findstr [0-9].\.)[0]).Split()[-1]
if ($IP -eq $null) {$IP_err_color = $errorcolor}
$FQDN = (Get-WmiObject win32_computersystem).DNSHostName+"."+(Get-WmiObject win32_computersystem).Domain
$LANAdaters = Get-NetAdapter -physical | where status -eq 'up'
if ($LANAdaters -eq $null) {$LANAdaters_err_color = " bgcolor='#F84A42'"}
$DNSServers = ((Get-DnsClientServerAddress -InterfaceIndex $LANAdaters.InterfaceIndex -AddressFamily IPv4 | select ServerAddresses).ServerAddresses | Out-String).replace("`r`n","<br>")
if ($DNSServers -eq $null) {$DNSServers_err_color = " bgcolor='#F84A42'"}
$LANAdaters = GetHTML ($LANAdaters | fl Name,InterfaceDescription,Status,LinkSpeed | Out-String)


$DefaultGateway_IP = (Get-wmiObject Win32_networkAdapterConfiguration | ?{$_.IPEnabled}).DefaultIPGateway
$DefaultGateway_PING = (Test-Connection (Get-NetRoute -DestinationPrefix 0.0.0.0/0 | Select-Object -ExpandProperty Nexthop) -Count 2 | Measure-Object -Property ResponseTime -Average).Average
if ($DefaultGateway_PING -eq $null) {$DefaultGateway_PING_err_color = " bgcolor='#F84A42'"}

$Exchange_IP = GetIP -ip_req $Exchange
if ($Exchange_IP -ne $null) {
    $Exchange_PING = ((Test-Connection $Exchange_IP -Count 2 | Measure-Object -Property ResponseTime -Average).Average | Out-String) + " ms"
    if ($Exchange_PING -eq " ms") {$Exchange_PING_err_color = $errorcolor}

    $Exchange_TraceRoute = GetHTML (Test-NetConnection $Exchange_IP -traceroute | fl | Out-String)
    if ($Exchange_TraceRoute -eq $null) {$Exchange_TraceRoute_err_color = $errorcolor}

    $Exchange_Port = GetPort GetPort -GetPort_ip $Exchange_IP -GetPort_port 443
    if ($Exchange_Port -eq ":443 - False")
    {
        $Exchange_Port_err_color = $errorcolor
        $Exchange_Port_err_skip = 1
    }

    #Skip web tests if 443 port not works.
    if ($Exchange_Port_err_skip -ne 1) {
        $Exchange_MAPIHC = GetHTML (GetURL -GetURL_req $Exchange_MAPIHC_URL | Out-String)
        $Exchange_EWSHC = GetHTML (GetURL -GetURL_req $Exchange_EWSHC_URL | Out-String)
        $Exchange_ADHC = GetHTML (GetURL -GetURL_req $Exchange_ADHC_URL | Out-String)
        $Exchange_EMSMDB = GetHTML (GetURL -GetURL_req $Exchange_EMSMDB_URL | Out-String)
        $Exchange_AD1 = GetURL -GetURL_req $Exchange_AD_URL1
        $Exchange_AD2 = GetURL -GetURL_req $Exchange_AD_URL2
    }

    if (($Exchange_MAPIHC -eq "") -or ($Exchange_MAPIHC -eq $null)) {$Exchange_MAPIHC_err_color = $errorcolor}
    if (($Exchange_EWSHC -eq "") -or ($Exchange_EWSHC -eq $null)) {$Exchange_EWSHC_err_color = $errorcolor}
    if (($Exchange_ADHC -eq "") -or ($Exchange_ADHC -eq $null)) {$Exchange_ADHC_err_color = $errorcolor}
    if (($Exchange_EMSMDB -eq "") -or ($Exchange_EMSMDB -eq $null)) {$Exchange_EMSMDB_err_color = $errorcolor}

    $Exchange_netstat = GetHTML (Get-NetTCPConnection -RemoteAddress $Exchange_IP -RemotePort 443 | Out-String)
    if (($Exchange_netstat -eq "") -or ($Exchange_netstat -eq $null)) {$Exchange_netstat_err_color = $errorcolor}
}
else
{
    $Exchange_IP = GetHTML ($error[0] | Out-String)
    $Exchange_IP_err_color = $errorcolor
    $Exchange_PING_err_color = $errorcolor
    $Exchange_TraceRoute_err_color = $errorcolor
    $Exchange_Port_err_color = $errorcolor
    $Exchange_MAPIHC_err_color = $errorcolor
    $Exchange_EWSHC_err_color = $errorcolor
    $Exchange_ADHC_err_color = $errorcolor
    $Exchange_EMSMDB_err_color = $errorcolor
    $Exchange_netstat_err_color = $errorcolor
}

$hosts_search = New-Object System.Collections.ArrayList
$hosts_search1 = (Select-String -Path $hosts_file -pattern $Exchange).line
$hosts_search2 = (Select-String -Path $hosts_file -pattern $SMTPDomain).line
$hosts_search3 = (Select-String -Path $hosts_file -pattern $Exchange_IP).line
$hosts_search = GetHTML(($hosts_search1.Split([Environment]::NewLine) + $hosts_search2.Split([Environment]::NewLine) + $hosts_search3.Split([Environment]::NewLine))| select -Unique | Out-String)

if ($hosts_search -eq $null) {$hosts_search = "The hosts file doesn't contains the Exchange Server addresses."}


#Outlook
$OutlookFiles = ""; $OutlookPath = ""; $version = "";
foreach ($OutlookPath in $OutlookPaths) {
    if([System.IO.File]::Exists($OutlookPath)){
        $version = (Get-ChildItem $OutlookPath -Recurse | Select-Object -ExpandProperty VersionInfo).FileVersion
        $OutlookFiles += "$OutlookPath | Version: $version"
    }
}
if ($OutlookFiles -eq "") {$OutlookFiles_err_color = $errorcolor}
$OutlookFiles = GetHTML($OutlookFiles)

$OutlookRegs = ""
foreach ($OutlookReg in $OutlookRegAD) {
    if (Test-Path $OutlookReg) {$OutlookRegs += (Get-ItemProperty $OutlookReg | Out-String)}
}
$OutlookRegs = GetHTML($OutlookRegs)


$OS = GetHTML (Get-CimInstance Win32_OperatingSystem | fl Caption, Version, ServicePackMajorVersion, OSArchitecture | Out-String)

$EventLogSearch = GetHTML (Get-EventLog -LogName Application -source Outlook -After (Get-Date).AddMinutes(-$EventLog_Minutes) | fl Time,EntryType,Source,Message | Out-String)


# HTML Generation
$Report = New-Object System.Collections.ArrayList
$Report.Add("<!DOCTYPE html><html><head><title>Outlook to Exchange Connection Test</title><style>html * { font-family: Calibri}; table, tr, td {border-collapse: collapse; border: 1px solid;}; .style1 {border-collapse: collapse; border: 0px solid;}</style></head><body>")
$Report.Add("<h1>$date2</h1>")
$Report.Add("<table border='1' cellpadding='0' cellspacing='0'")

$Report.Add("<tr><td colspan=2 bgcolor='#69B7C8' style='padding: 5px;' align=center><b>Workstation</b></td></tr>")
$Report.Add("<tr><td bgcolor='#7FCEDF' style='padding: 5px;'>Computer name</td><td style='padding: 5px;'>$FQDN</td></tr>")
$Report.Add("<tr><td bgcolor='#7FCEDF' style='padding: 5px;'>OS</td><td style='padding: 5px;'>$OS</td></tr>")
$Report.Add("<tr><td bgcolor='#7FCEDF' style='padding: 5px;'>IP</td><td style='padding: 5px;'$IP_err_color>$IP</td></tr>")
$Report.Add("<tr><td bgcolor='#7FCEDF' style='padding: 5px;'>Default Gateway</td><td style='padding: 5px;'$DefaultGateway_PING_err_color>IP = $DefaultGateway_IP <br>PING = $DefaultGateway_PING ms</td></tr>")
$Report.Add("<tr><td bgcolor='#7FCEDF' style='padding: 5px;'>Physical Adapters</td><td style='padding: 5px;'$LANAdaters_err_color>$LANAdaters</td></tr>")
$Report.Add("<tr><td bgcolor='#7FCEDF' style='padding: 5px;'>DNS Servers</td><td style='padding: 5px;'$DNSServers_err_color>$DNSServers</td></tr>")
$Report.Add("<tr><td bgcolor='#7FCEDF' style='padding: 5px;'>hosts</td><td style='padding: 5px;'>$hosts_search</td></tr>")


$Report.Add("<tr><td colspan=2 bgcolor='#69B7C8' style='padding: 5px;' align=center><b>Exchange</b></td></tr>")
$Report.Add("<tr><td bgcolor='#7FCEDF' style='padding: 5px;'>Exchange Server</td><td style='padding: 5px;'>$Exchange</td></tr>")
$Report.Add("<tr><td bgcolor='#7FCEDF' style='padding: 5px;'>IP</td><td style='padding: 5px;'$Exchange_IP_err_color>$Exchange_IP</td></tr>")
$Report.Add("<tr><td bgcolor='#7FCEDF' style='padding: 5px;'>Ping</td><td style='padding: 5px;'$Exchange_PING_err_color>$Exchange_PING</td></tr>")
$Report.Add("<tr><td bgcolor='#7FCEDF' style='padding: 5px;'>Telnet</td><td style='padding: 5px;'$Exchange_Port_err_color>$Exchange_Port</td></tr>")
$Report.Add("<tr><td bgcolor='#7FCEDF' style='padding: 5px;'>Traceroute</td><td style='padding: 5px;'$Exchange_TraceRoute_err_color>$Exchange_TraceRoute</td></tr>")
$Report.Add("<tr><td bgcolor='#7FCEDF' style='padding: 5px;'>MAPI Health Check</td><td style='padding: 5px;'$Exchange_MAPIHC_err_color>$Exchange_MAPIHC</td></tr>")
$Report.Add("<tr><td bgcolor='#7FCEDF' style='padding: 5px;'>EWS Health Check</td><td style='padding: 5px;'$Exchange_EWSHC_err_color>$Exchange_EWSHC</td></tr>")
$Report.Add("<tr><td bgcolor='#7FCEDF' style='padding: 5px;'>Autodiscover Health Check</td><td style='padding: 5px;'$Exchange_ADHC_err_color>$Exchange_ADHC</td></tr>")
$Report.Add("<tr><td bgcolor='#7FCEDF' style='padding: 5px;'>MAPI EMSMDB</td><td style='padding: 5px;'$Exchange_EMSMDB_err_color>$Exchange_EMSMDB</td></tr>")
$Report.Add("<tr><td bgcolor='#7FCEDF' style='padding: 5px;'>Open sessions to Exchange</td><td style='padding: 5px;'$Exchange_netstat_err_color>$Exchange_netstat</td></tr>")


$Report.Add("<tr><td colspan=2 bgcolor='#69B7C8' style='padding: 5px;' align=center><b>Outlook</b></td></tr>")
$Report.Add("<tr><td bgcolor='#7FCEDF' style='padding: 5px;'>Founded versions</td><td style='padding: 5px;'$OutlookFiles_err_color>$OutlookFiles</td></tr>")
$Report.Add("<tr><td bgcolor='#7FCEDF' style='padding: 5px;'>Autodiscover settings in registry</td><td style='padding: 5px;'>$OutlookRegs</td></tr>")
$Report.Add("<tr><td bgcolor='#7FCEDF' style='padding: 5px;'>Event Log (-$EventLog_Minutes minutes)</td><td style='padding: 5px;'>$EventLogSearch</td></tr>")

$Report.Add("</table></body></html>")

Add-Content $HTMLReportFile $Report
Invoke-Item $HTMLReportFile