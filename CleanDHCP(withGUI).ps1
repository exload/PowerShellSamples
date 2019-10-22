function Show-DHCP {
[void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
[void][reflection.assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
[void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[System.Windows.Forms.Application]::EnableVisualStyles()
$DHCPFree = New-Object 'System.Windows.Forms.Form'
$ClearSelected = New-Object 'System.Windows.Forms.Button'
$IPAddresses = New-Object 'System.Windows.Forms.CheckedListBox'
$GetLease = New-Object 'System.Windows.Forms.Button'
$GetScopeLists = New-Object 'System.Windows.Forms.Button'
$ScopeLists = New-Object 'System.Windows.Forms.ComboBox'
$DHCPServers = New-Object 'System.Windows.Forms.ComboBox'
$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
$DHCPFree_Load={Get-DhcpServerInDC | where { $_.DnsName -ne "" } | select DnsName | %{ [void]$DHCPServers.Items.Add($_.DnsName)} }
 
$GetScopeLists_Click={
$DHCPFree.Enabled = $false
[void]$ScopeLists.Items.Clear()
$ScopeLists.Text = ""
[void]$IPAddresses.Items.Clear()
if (Test-Connection $DHCPServers.Text -Count 2 -ErrorAction SilentlyContinue){Get-DhcpServerv4Scope -ComputerName $DHCPServers.Text -ErrorAction SilentlyContinue | %{ [void]$ScopeLists.Items.Add($_.ScopeId) } }
$DHCPFree.Enabled = $true
}
 
$GetLease_Click={
$DHCPFree.Enabled = $false
$IPAddresses.Items.Clear()
$AllRecords = Get-DhcpServerv4Lease -ScopeId $ScopeLists.Text -ComputerName $DHCPServers.Text
foreach ($Record in $AllRecords){
if (($Record.AddressState -ne "InactiveReservation") -and ($Record.AddressState -ne "ActiveReservation")){
if (!(Test-Connection $Record.IPAddress -Count 2 -ErrorAction SilentlyContinue)){$IPAddresses.Items.Add($Record.IPAddress)}
}
}
$DHCPFree.Enabled = $true
}
 
$ClearSelected_Click = {
$DHCPFree.Enabled = $false
foreach ($IP in $IPAddresses.CheckedItems){
write-host $IP
Remove-DhcpServerv4Lease -IPAddress $IP -ComputerName $DHCPServers.Text
}
[void]$ScopeLists.Items.Clear()
$ScopeLists.Text = ""
[void]$IPAddresses.Items.Clear()
$DHCPFree.Enabled = $true
}
$Form_StateCorrection_Load=
{
$DHCPFree.WindowState = $InitialFormWindowState
}
 
$Form_Cleanup_FormClosed=
{
try
{
$ClearSelected.remove_Click($ClearSelected_Click)
$GetLease.remove_Click($GetLease_Click)
$GetScopeLists.remove_Click($GetScopeLists_Click)
$DHCPFree.remove_Load($DHCPFree_Load)
$DHCPFree.remove_Load($Form_StateCorrection_Load)
$DHCPFree.remove_FormClosed($Form_Cleanup_FormClosed)
}
catch { Out-Null }
}
$DHCPFree.SuspendLayout()
$DHCPFree.Controls.Add($ClearSelected)
$DHCPFree.Controls.Add($IPAddresses)
$DHCPFree.Controls.Add($GetLease)
$DHCPFree.Controls.Add($GetScopeLists)
$DHCPFree.Controls.Add($ScopeLists)
$DHCPFree.Controls.Add($DHCPServers)
$DHCPFree.AutoScaleDimensions = '6, 13'
$DHCPFree.AutoScaleMode = 'Font'
$DHCPFree.ClientSize = '280, 446'
$DHCPFree.MaximizeBox = $False
$DHCPFree.MinimizeBox = $False
$DHCPFree.Name = 'DHCPFree'
$DHCPFree.ShowIcon = $False
$DHCPFree.add_Load($DHCPFree_Load)
$ClearSelected.Location = '53, 417'
$ClearSelected.Name = 'ClearSelected'
$ClearSelected.Size = '192, 26'
$ClearSelected.TabIndex = 5
$ClearSelected.Text = 'Free selected address'
$ClearSelected.UseVisualStyleBackColor = $True
$ClearSelected.add_Click($ClearSelected_Click)
$IPAddresses.FormattingEnabled = $True
$IPAddresses.Location = '9, 124'
$IPAddresses.Name = 'IPAddresses'
$IPAddresses.Size = '259, 289'
$IPAddresses.TabIndex = 4
$GetLease.Location = '86, 92'
$GetLease.Name = 'GetLease'
$GetLease.Size = '122, 28'
$GetLease.TabIndex = 3
$GetLease.Text = 'Get lease'
$GetLease.UseVisualStyleBackColor = $True
$GetLease.add_Click($GetLease_Click)
$GetScopeLists.Location = '84, 35'
$GetScopeLists.Name = 'GetScopeLists'
$GetScopeLists.Size = '125, 27'
$GetScopeLists.TabIndex = 2
$GetScopeLists.Text = 'Get scope lists'
$GetScopeLists.UseVisualStyleBackColor = $True
$GetScopeLists.add_Click($GetScopeLists_Click)
$ScopeLists.FormattingEnabled = $True
$ScopeLists.Location = '11, 67'
$ScopeLists.Name = 'ScopeLists'
$ScopeLists.Size = '258, 21'
$ScopeLists.TabIndex = 1
$DHCPServers.FormattingEnabled = $True
$DHCPServers.Location = '11, 10'
$DHCPServers.Name = 'DHCPServers'
$DHCPServers.Size = '258, 21'
$DHCPServers.TabIndex = 0
$DHCPFree.ResumeLayout()
$InitialFormWindowState = $DHCPFree.WindowState
$DHCPFree.add_Load($Form_StateCorrection_Load)
$DHCPFree.add_FormClosed($Form_Cleanup_FormClosed)
return $DHCPFree.ShowDialog()
} 
Show-DHCP