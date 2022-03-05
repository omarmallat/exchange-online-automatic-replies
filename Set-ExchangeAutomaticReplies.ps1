[OutputType([string])]
param ([Parameter(Mandatory=$true)]
        [string]$credentialName,
        [Parameter(Mandatory=$true)]
        [string]$mailSource,
        [Parameter(Mandatory=$true)]
        [string]$mailTarget,
        [Parameter(Mandatory=$true)]
        [string]$mailBody)

Connect-AzureAD -Credential (Get-AutomationPSCredential -Name $credentialName) | Out-Null
Connect-ExchangeOnline -Credential (Get-AutomationPSCredential -Name $credentialName) -CommandName 'Get-MailboxFolderPermission','Get-MailboxFolderStatistics','Set-MailboxAutoReplyConfiguration' | Out-Null

#region: Input
if ($mailTarget.Contains('@'))
{
    $checkEmail = $true
    # Check manager
    if ((@((Get-AzureADUser -Filter "Mail eq '$mailTarget'" | Get-AzureADUserManager).UserPrincipalName) | Select-Object -Unique) -contains $mailSource)
    {
        $checkSuperordinate = $true
    }
    else
    {
        $checkSuperordinate = $false
    }
    # Check coworker
    if ((@((Get-AzureADUser -Filter "Mail eq '$mailTarget'" | Get-AzureADUserManager | Get-AzureADUserDirectReport).UserPrincipalName) | Select-Object -Unique) -contains $mailSource)
    {
        $checkCoordinate = $true
    }
    else
    {
        $checkCoordinate = $false
    }
    # Check direct reports
    if ((@((Get-AzureADUser -Filter "Mail eq '$mailTarget'" | Get-AzureADUserDirectReport).UserPrincipalName) | Select-Object -Unique) -contains $mailSource)
    {
        $checkSubordinate = $true
    }
    else
    {
        $checkSubordinate = $false
    }
    # Check editors
    if ((Get-AzureADUser -Filter "UserPrincipalName eq '$mailSource'").DisplayName -in ((@((Get-MailboxFolderStatistics -Identity $mailTarget -FolderScope Calendar | Where-Object{$_.FolderType -eq 'Calendar'}), (Get-MailboxFolderStatistics -Identity $mailTarget -FolderScope Inbox | Where-Object{$_.FolderType -eq 'Inbox'})) | ForEach-Object{Get-MailboxFolderPermission $_.Identity.replace('\',':\')} | Where-Object{$_.AccessRights -in @('Editor','PublishingEditor','Owner')}).User.DisplayName | Select-Object -Unique))
    {
        $checkPermission = $true
    }
    else
    {
        $checkPermission = $false
    }
    # Check results
    if ($checkSuperordinate -or $checkCoordinate -or $checkSubordinate -or $checkPermission)
    {
        $text = "<html><body><div style='font-size: 11pt; font-family: Arial'>$mailBody</div></body></html>"
        Set-MailboxAutoReplyConfiguration -Identity $mailTarget -AutoReplyState 'Enabled' -InternalMessage $text -ExternalMessage $text -ExternalAudience 'All'
    }
}
else
{
    $checkEmail = $false
}
#endregion

#region: Output
$output = [System.Text.StringBuilder]::new()
$output.AppendLine(@"
<div>
<style>
table, th, td {
border: none;
border-collapse: collapse;
}
th, td {
padding: 5px;
text-align: left;
vertical-align: top;
}
.green {
border-left: 4pt solid darkgreen;
padding-left: 4pt;
background-color: lightgreen
}
.yellow {
border-left: 4pt solid darkgoldenrod;
padding-left: 4pt;
background-color: lightgoldenrodyellow
}
.red {
border-left: 4pt solid darkred;
padding-left: 4pt;
background-color: lightcoral
}
</style>
<table>
<tr>
<th>Check</th>
<th>Result</th>
</tr>
<tr>
<td>Requester provided an email address</td>
"@) | Out-Null
if ($checkEmail)
{
    $output.AppendLine('<td class="green">Yes</td>') | Out-Null
}
else
{
    $output.AppendLine('<td class="red">No</td>') | Out-Null
}
$output.AppendLine(@"
</tr>
<tr>
<td>Requester is the manager</td>
"@) | Out-Null
if ($checkEmail)
{
    if ($checkSuperordinate)
    {
        $output.AppendLine('<td class="green">Yes</td>') | Out-Null
    }
    else
    {
        $output.AppendLine('<td class="yellow">No</td>') | Out-Null
    }
}
else
{
    $output.AppendLine('<td class="yellow">Skipped</td>') | Out-Null
}
$output.AppendLine(@"
</tr>
<tr>
<td>Requester is a coworker</td>
"@) | Out-Null
if ($checkEmail)
{
    if ($checkCoordinate)
    {
        $output.AppendLine('<td class="green">Yes</td>') | Out-Null
    }
    else
    {
        $output.AppendLine('<td class="yellow">No</td>') | Out-Null
    }
}
else
{
    $output.AppendLine('<td class="yellow">Skipped</td>') | Out-Null
}
$output.AppendLine(@"
</tr>
<tr>
<td>Requester is a direct report</td>
"@) | Out-Null
if ($checkEmail)
{
    if ($checkSubordinate)
    {
        $output.AppendLine('<td class="green">Yes</td>') | Out-Null
    }
    else
    {
        $output.AppendLine('<td class="yellow">No</td>') | Out-Null
    }
}
else
{
    $output.AppendLine('<td class="yellow">Skipped</td>') | Out-Null
}
$output.AppendLine(@"
</tr>
<tr>
<td>Requester has inbox/calendar permission</td>
"@) | Out-Null
if ($checkEmail)
{
    if ($checkPermission)
    {
        $output.AppendLine('<td class="green">Yes</td>') | Out-Null
    }
    else
    {
        $output.AppendLine('<td class="yellow">No</td>') | Out-Null
    }
}
else
{
    $output.AppendLine('<td class="yellow">Skipped</td>') | Out-Null
}
$output.AppendLine(@"
</tr>
<tr>
<td>Configure automatic replies</td>
"@) | Out-Null
if ($checkEmail)
{
    if ($checkSuperordinate -or $checkCoordinate -or $checkSubordinate -or $checkPermission)
    {
        $output.AppendLine('<td class="green">Granted</td>') | Out-Null
    }
    else
    {
        $output.AppendLine('<td class="red">Denied</td>') | Out-Null
    }
}
else
{
    $output.AppendLine('<td class="yellow">Skipped</td>') | Out-Null
}
$output.AppendLine(@"
</tr>
</table>
</div>
"@) | Out-Null
$output.ToString() | Write-Output
#endregion
