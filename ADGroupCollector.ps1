<#	
    .NOTES
    ===========================================================================
    Created with: 	VS
    Created on:   	03/02/2022
    Created by:   	Chris Healey
    Organization: 
    Version:        0.1	
    Filename:       ADGroupCollector.ps1
    Project path:   https://github.com/healeychris/ADGroupCollector
    
    ===========================================================================
    .DESCRIPTION
    This script will collect all Distribution group types from AD/Exchange and produce a report details information found
    This script can be used to identify owners and number of members per group, plus other details
    .NOTES

    How to Run: -
    .\ADGroupCollector.ps1

    Expectation is that a connection to Exchange Powshell and  ActiveDirectory has been established in order to run the script.


#>

####### Variable list #######
$Version                                                = "0.1"                                                       # Version of script
Clear-Host                                                                                                            # Clear screen
$host.ui.RawUI.WindowTitle                              = 'Distribution List Collector'                               # Title for Status Bar
$ExportFile                                             = "GroupReport_$((get-date).ToString('yyyyMMdd_HHmm')).csv"   # Export File name

##############################
cls

Write-host "Getting Exchange Distribution Groups, Please Wait...." -ForegroundColor GREEN

# Collect Distribution Groups
$GroupDataRaw = Get-DistributionGroup -resultsize unlimited | Select-Object SamAccountName,Displayname,Alias,GroupType,MemberJoinRestriction,MemberDepartRestriction,OrganizationalUnit,CustomAttribute5,HiddenFromAddressListsEnabled,PrimarySMTPAddress,RecipientType,RecipientTypeDetails,WhenCreated,WhenChanged,DistinguishedName,Description,Info,ManagedBy	

# Collect and build data for export
Foreach ($Group in $GroupDataRaw) {

    # Build simple string
    $GroupDisplayName = $Group.Displayname

    # Create counter progress monitor
    $counter++
    Write-Progress -Activity 'Processing Groups' -CurrentOperation $GroupDisplayName -PercentComplete (($counter / $GroupDataRaw.count) * 100)

    # Get Base group members and memberof
    $GroupADDetails = Get-ADGroup -Identity $Group.distinguishedName -Properties Members,MemberOf,Description,info
    
    # Get Member Count
    $GroupMembersCount = $GroupADDetails.members | Measure-Object | Select-Object -ExpandProperty Count

    # Get Memberof Count
    $GroupMembersOfCount = $GroupADDetails.MemberOf | Measure-Object | Select-Object -ExpandProperty Count

    # Get owners details
    $GroupOwnerDetails = ''
    Foreach ($Owner in $Group.ManagedBy.DistinguishedName){
    
        try {$GroupOwnerMail = Get-ADUser -identity $Owner -properties Mail -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Mail}
        Catch {Write-host "$Owner was not found from $GroupDisplayName" -ForegroundColor YELLOW}

        
        $GroupOwnerDetails += $GroupOwnerMail + -join ";"      

   
    }
    if (!($GroupOwnerDetails)) {$GroupOwnerDetails = 'No Owner Details Found'}


    # Create PSobject for Group Data
    $GroupData = @()
    $GroupData = [pscustomobject][ordered]@{

    'LoginID'                         = $Group.SAmaccountName
    'Displayname'                     = $Group.Displayname
    'Alias'                           = $Group.Alias
    'GroupType'                       = $Group.GroupType
    'MemberJoinRestriction'           = $Group.MemberJoinRestriction
    'MemberDepartRestriction'         = $Group.MemberDepartRestriction
    'OrganizationalUnit'              = $Group.OrganizationalUnit
    'CustomAttribute5'                = $Group.CustomAttribute5
    'HiddenFromAddressListsEnabled'   = $Group.HiddenFromAddressListsEnabled
    'PrimarySMTPAddress'              = $Group.PrimarySMTPAddress
    'RecipientType'                   = $Group.RecipientType
    'RecipientTypeDetails'            = $Group.RecipientTypeDetails
    'WhenCreated'                     = $Group.WhenCreated
    'WhenChanged'                     = $Group.WhenChanged
    'DistinguishedName'               = $Group.DistinguishedName
    'Description'                     = $GroupADDetails.Description
    'Info'                            = $GroupADDetails.Info
    'NumberofMembers'                 = $GroupMembersCount 
    'NumberofMembersOf'               = $GroupMembersOfCount 
    'ManagedBy'                       = $GroupOwnerDetails
    }


    # Export data to CSV File
    $GroupData| Export-Csv $ExportFile -NoTypeInformation -Append

}
