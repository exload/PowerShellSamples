<#
	.SYNOPSIS
		Script to parse Tiny proxy log into an HTML-based report.

	.DESCRIPTION
		The script copies the Tiny server proxy log and parses into an HTML-based report. It may also be called by another script.
	
	.NOTES
		Requires powershell module "PoSH-SSH" version "2.2" to be installed from gallery.technet.microsoft.com
        From PoSH-SSH use only this cmdlets: New-SSHTrustedHost, Get-SCPFile, Get-SSHTrustedHost.
        To access the proxy server via SSH, the script asks for user credentials.

	.PARAMETER RunFromScript
		The parameter value indicates how script was invoked: by user or by another script.
		Valid values: 0 - was called by user, 1 - launched from another script
		Default value: 0
		
	.PARAMETER UTCTime
		Sets the timestamp with which records will be searched.
		Default value: '00:00:01'
	
	.PARAMETER ClientIP
		Sets the IP address of the computer where the software was tested
		Default value: '%some_ip%'
	
    .PARAMETER ProxyIP
		Sets the IP address of the proxy server
		Default value: '%proxy_ip%'

    .PARAMETER FingerPrint
		Sets the fingerprint of the proxy server certificate. Must be kept up to date.
        Default value is valid on 28/08/19
		Default value: '%cert_thumbprint%'
	
    .PARAMETER ProxyLog
		Path to the log file of the proxy server.
		Default value: '/var/log/tinyproxy/tinyproxy.log'

    .PARAMETER LocalFile
		The path where to copy the log file from the proxy server.
		Default value: 'C:\Distr\tiny.log'

	.PARAMETER ProxyReportPath
		Path to the HTML report file.
		Default value: 'C:\Distr\ProxyReport.html'

	.EXAMPLE
		Example of how to run the script:
		powershell -NoProfile -executionpolicy bypass .\ReadProxyLog.ps1 -ProxyIP '10.0.0.1'
		powershell -NoProfile -executionpolicy bypass .\ReadProxyLog.ps1 -ProxyReportPath 'C:\Reports\ProxyReport.html'
#>

#Requires -Modules @{ ModuleName="PoSH-SSH"; RequiredVersion="2.2" }

[CmdletBinding()]
param
(
    [parameter(mandatory=$false)][ValidateSet(0,1)][int]$RunFromScript=0,
    [parameter(mandatory=$false)][string]$UTCTime='00:00:01',
    [parameter(mandatory=$false)][string]$ClientIP = '%some_ip%',
	[parameter(mandatory=$false)][string]$ProxyIP = '%proxy_ip%',
	[parameter(mandatory=$false)][string]$FingerPrint = '%cert_thumbprint%',
    [parameter(mandatory=$false)][string]$ProxyLog = '/var/log/tinyproxy/tinyproxy.log',
    [parameter(mandatory=$false)][string]$LocalFile = 'C:\Distr\tiny.log',
    [parameter(mandatory=$false)][string]$ProxyReportPath='C:\Distr\ProxyReport.html'
)

New-SSHTrustedHost -SSHHost $ProxyIP -FingerPrint $FingerPrint
Get-SCPFile -LocalFile $LocalFile -RemoteFile $Proxy_log -ComputerName $ProxyIP -Credential $(Get-Credential -Message 'Need credentials for access to proxy server')
Get-SSHTrustedHost | Remove-SSHTrustedHost

$Content = Get-Content $LocalFile
$ContentLength = $Content.Count

Remove-Item -Path $LocalFile -Force -Verbose

$HTML = 'Finding entries after ' + $UTCTime + '<br>'
for ($i = 0; $i -lt $ContentLength; $i++)
{
    $Entry = $Content[$i].split()
    if ($Entry[0] -eq 'CONNECT' -and ($Entry[11] -eq $ClientIP))
    {
        $UTCProxy = $Entry[5]
        if ($UTCProxy -ge $UTCTime)
        {
            $HTML += '' + ($i + 1) + ' ' + $Content[$i] + '<br>'
            if (($i + 1) -lt $ContentLength)
            {
                $HTML += '' + ($i + 2) + ' ' + $Content[$i+1] + '<br>'
            }
        }
    }
}
if ($RunFromScript -eq 0)
{
    ConvertTo-Html -PostContent $HTML | Out-File -FilePath $ProxyReportPath
}
else
{
    return $HTML
}