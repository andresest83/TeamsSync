# TeamsSync
Script to Sync on-premises Active Directory Security Groups with Microsoft Teams(which are just a type of Office 365 group). This script started as a project to sync a few teams based on https://gallery.technet.microsoft.com/Sync-AD-Group-with-Teams-74598786 and as I put it into production in a company I worked for I started adding more features as they appeared.

# Architecture
The architecture is pretty simple. An Active Directory Group is linked to a MS Teams by inserting the Office 365 Group Id of the MS Team in its `extensionAttribute5`. The script looks for all the Active Directory Security Groups in a particular Organizational Unit that have something set in their extensionAttribute5 and then adds/removes users from the corresponding MS Team based on the Active Directory Security Group Members.

# Prerequisites
`
*Install-PackageProvider -Name nuget -MinimumVersion 2.8.5.201 -Force -Scope AllUsers
*Install-Module microsoftteams -Scope AllUsers -Force
`

# Setup
Please setup the following variables accordingly. They are at the beggining of the script.
`
*server - The script targets a specific domain controller, therefore avoding replication delays.
*myOU - The OU where your Security Groups are.
*sendMailFrom - Account used to send notifications
*sendMailTo - Account to send notifications to
*teamsCredentials - An account with read/write permissions on Office 365/Teams and read permissions on Active Directory.
`
