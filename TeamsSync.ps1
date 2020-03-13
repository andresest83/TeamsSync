# Make sure to install the Nuget Package Manager, the Teams Module and the AD Module
# before running this script. Line 4 and 5 show you how

# Install-PackageProvider -Name nuget -MinimumVersion 2.8.5.201 -Force -Scope AllUsers
# Install-Module microsoftteams -Scope AllUsers -Force

# Variables declaration - Please change it accordingly

[string]$server = "dc1.contoso.com"
$myOU = "OU=Groups,DC=contoso,DC=com"
$sendMailFrom = "teams.manager@contoso.com"
$sendMailTo = "blabla.contoso.onmicrosoft.com@emea.teams.ms" # The script assumes you send the notifications to a channel in teams

# Credentials for both Teams and Active Directory (The script assumes it is the same account for Office 365
# and Active Directory on-premises in the format admin@contoso.com

$teamsCredentials = Get-Credential

# This function will allow the script to work with nested groups

function NestedGroup
{
  param([string]$group_member)
  $groupuserinf = [System.Collections.ArrayList]@()

  $members = get-adgroup -Server $server $group_member | get-adgroupmember -Server $server
  #$members = Get-ADGroup -Server $server -Filter {name -eq $group_member} | get-adgroupmember -Server $server

  foreach ($member in $members)
  {
    if ($member.objectClass -eq "group")
    {
      NestedGroup ($member)
    }

    if ($member.objectClass -eq "user")
    {
      $users = get-aduser -Server $server $member | Where-Object { ($_.UserPrincipalName.SubString(0,2) -ne "a-") -and ($_.Enabled -eq "True") } | Select-Object UserPrincipalName | ForEach-Object { $groupuserinf.Add($_) }
    }
  }
  return $groupuserinf
}

# This function allows the script to detect circular groups

function Get-CircularNestedGroups {
  $groups = @()
  $groups = get-adgroup -Server $server -SearchBase $myOU -Filter { GroupCategory -EQ "Security" }
  $return = @()

  # Parse through groups and add further encapsulated groups
  while ($groups) {
    foreach ($group in $groups) {
      $done = @()
      [array]$tmp_grps = (get-adgroup -Server $server $group -Properties memberOf).memberOf
      [array]$grps += $tmp_grps

      # Recursive parsing through the submembergroups
      while ($tmp_grps) {
        foreach ($tmp_g in $tmp_grps) {
          [array]$done += $tmp_g
          [array]$sub_grps = (get-adgroup -Server $server $tmp_g -Properties memberOf).memberOf

          # Add submembergroups to temporary array
          if ($sub_grps) {
            $tmp_grps += $sub_grps
            $grps += $sub_grps
          }

          # Remove already parsed groups from temporary array
          $tmp_grps = $tmp_grps | Where-Object { $_ -ne $tmp_g -and $done -notcontains $_ }
          $sub_grps = ''
        }

      }

      # Add circular nested groups to return value
      if ($grps -contains $group) {
        $return += $group
      }

      # Remove already parsed groups from array
      $groups = $groups | Where-Object { $_ -ne $group }

      # Clean up
      Remove-Variable -Name tmp_grps,done,grps,Group-Object
    }
  }

  return $return
}

# This function manages all logging for the scripts
# Make sure you have a folder named "C:\Logs" or change this accordingly

$Logfile = "C:\Logs\Teams_Sync_$(Get-Date -Format dd-MM-yyyy).log"

function LogWrite {
  param([string]$logstring)
  $time = $(Get-Date -DisplayHint Time).ToString("HH:mm:ss")
  Add-Content $Logfile -Value "[$time] $logstring "
}

# Prechek of circular groups, if not passed the script execution is cancelled
$circularGroups = @()
$circularGroups = Get-CircularNestedGroups | Select-Object samAccountName,ObjectGUID
$circularGroupsHTML = $circularGroups | ConvertTo-Html -As Table -Head " " -Title " " -PreContent " "

if (-not [string]::IsNullOrEmpty($circularGroups))
{

  $body = @"
<br>
The following groups are circulary nested. Please fix this issue to resume the Teams sync.<br>
Please click on "See more" to see the full list.<br>
$circularGroupsHTML
"@
  $mailprops = @{
    From = '$sendMailFrom'
    To = '$sendMailTo'
    SmtpServer = 'smtp.office365.com'
    BodyAsHtml = $true
    Subject = 'Circular Groups Issue'
  }
  Send-MailMessage @mailprops -Body $body -Credential $teamsCredentials -UseSSL
  break
}


# Let's now retrieve the name of the AD Group and the MS Team using Extension Attribute 5 in Active Directory #

$Sam_Ea5_All = get-adgroup -Server $server -SearchBase $myOU -Filter { GroupCategory -EQ "Security" } -Properties Name,extensionAttribute5
$Sam_Ea5 = @()
$MS_Team = @()
$AD_Group = @()

foreach ($group in $Sam_Ea5_All)
{
  if (-not [string]::IsNullOrEmpty($group.extensionAttribute5))
  { $Sam_Ea5 += $group }
}


foreach ($item in $Sam_Ea5)
{ $MS_Team += $item.extensionAttribute5
  $AD_Group += $item.Name
}

# This part of the script does the actual sync

Connect-MicrosoftTeams -Credential $teamsCredentials

for ($i = 0; $i -lt $AD_Group.length; $i++) {
  $AD_Group[$i]
  $groupuserinfo = NestedGroup ($AD_Group[$i])
  $groupuserinfo = $groupuserinfo | Sort-Object -Property UserPrincipalName -Unique
  $errorMsgTeamUserInfo = $null
  $ErrorMessage = $null

  try {
    $teamuserinfo = get-teamuser -groupid $MS_Team[$i] | Sort-Object -Property user -Unique
  }
  catch {
    $errorMsgTeamUserInfo = $_.Exception.Message
  }

  if ($errorMsgTeamUserInfo -ne $null) {
    LogWrite ($errorMsgTeamUserInfo)
    Send-MailMessage -From $sendMailFrom -To $sendMailTo -Subject "Error getting MS Team info." -SmtpServer smtp.office365.com -Body $errorMsgTeamUserInfo -Credential $teamsCredentials -UseSSL -Port 587 -Encoding ([System.Text.Encoding]::UTF8)

  }
  else
  {
    if (($teamuserinfo.length -eq 1) -and ($groupuserinfo.length -eq 0)) {
      $logMsg = "Please add members to the group: " + $AD_Group[$i]
      LogWrite ($logMsg)
    }
    else
    {
      if (($teamuserinfo.length -gt 1) -and ($groupuserinfo.length -eq 0)) {
        $teamonlyusers = @()

        foreach ($teamuser in $teamuserinfo) {
          if ($teamuser.User -ne "$sendMailFrom") {
            $teamonlyusers += $teamuser.User
          }
        }

        foreach ($teamonlyuser in $teamonlyusers) {
          $logMsg = "Removing User $teamonlyuser from Team: " + $AD_Group[$i]
          LogWrite ($logMsg)

          try {
            remove-teamuser -groupid $MS_Team[$i] -user $teamonlyuser
          }
          catch {
            $ErrorMessage = $_.Exception.Message
            LogWrite ($ErrorMessage)
            Send-MailMessage -From $sendMailFrom -To $sendMailTo -Subject "Error Removing User from Teams" -SmtpServer smtp.office365.com -Body $ErrorMessage -Credential $teamsCredentials -UseSSL -Port 587 -Encoding ([System.Text.Encoding]::UTF8)
          }
        }

      }
      else
      {
        $directoryonlyusers = Compare-Object -DifferenceObject $groupuserinfo.UserPrincipalName -ReferenceObject $teamuserinfo.User -IncludeEqual -ErrorAction silentlycontinue | Where-Object { $_.SideIndicator -eq "=>" }
        $directoryonlyusers = $directoryonlyusers | Sort-Object -Property InputObject -Unique
        foreach ($directoryonlyuser in $directoryonlyusers) {
          $adonlyuser = $directoryonlyuser.InputObject
          $logMsg = "Adding User $adonlyuser to Team: " + $AD_Group[$i]
          LogWrite ($logMsg)

          try {
            add-teamuser -groupid $MS_Team[$i] -user $adonlyuser
          }
          catch {
            $ErrorMessage = $_.Exception.Message
            LogWrite ($ErrorMessage)
            Send-MailMessage -From $sendMailFrom -To $sendMailTo -Subject "Error Adding User to MS Teams" -SmtpServer smtp.office365.com -Body $ErrorMessage -Credential $teamsCredentials -UseSSL -Port 587
          }
        }

        $teamonlyusers = Compare-Object -DifferenceObject $groupuserinfo.UserPrincipalName -ReferenceObject $teamuserinfo.User -IncludeEqual -ErrorAction silentlycontinue | Where-Object { ($_.SideIndicator -eq "<=") -and ($_.InputObject.SubString(0,2) -ne "a-") -and ($_.InputObject -ne "$sendMailFrom") }
        foreach ($teamonlyuser in $teamonlyusers) {

          $msteamonlyuser = $teamonlyuser.InputObject
          $logMsg = "Removing User $msteamonlyuser from Team: " + $AD_Group[$i]
          LogWrite ($logMsg)

          try {
            remove-teamuser -groupid $MS_Team[$i] -user $msteamonlyuser
          }
          catch {
            $ErrorMessage = $_.Exception.Message
            LogWrite ($ErrorMessage)
            Send-MailMessage -From $sendMailFrom -To $sendMailTo -Subject "Error Removing User from Teams" -SmtpServer smtp.office365.com -Body $ErrorMessage -Credential $teamsCredentials -UseSSL -Port 587 -Encoding ([System.Text.Encoding]::UTF8)
          }
        }
      }
    }
  }
}

Disconnect-MicrosoftTeams
