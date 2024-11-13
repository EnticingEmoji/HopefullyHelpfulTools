function Connect-Sharepoint {

    #  Connects to SharePoint Online for OneDrive admin:
    Write-Host
    Write-Host "Connecting to sharepoint..."
#Replace yourTenantHere with your tenant. EG: contoso.com
    $tenant = "yourTenantHere"
    Connect-SPOService -Url https://$tenant-admin.sharepoint.com
}

Connect-Sharepoint

$userName = Read-Host "What is the SamAccountName or Email address of the user?"

# Reassigns OneDrive access to members of ITOM_SN for file backup

#Replace yourTenantHere with your tenant. EG: contoso.com
$tenant = "yourTenantHere"
$upn = (Get-ADUser $userName -Properties UserPrincipalName).UserPrincipalName
$logonName = $upn.Split("@")[0]
$domain = $upn.Split("@")[1]
$domain = $domain.Split(".")[0]
$first = $logonName.Split(".")[0]
$last = $logonName.Split(".")[1]
$url = "$first"+"_"+"$last"+"_"+"$domain"+"_"+"com"
$url2 = "https://$tenant-my.sharepoint.com/personal/$url"



Write-Host
#Replace "securityGroupContainingadminstoAdd" with a security group that contains the members you want to add as a share point admin.
$owners = Get-ADGroupMember "securityGroupContainingadminstoAdd"
$owners | foreach {
    $newOwner = (Get-ADUser $_ -Properties UserPrincipalName).UserPrincipalName
    Set-SPOUser -Site $url2 -LoginName $newOwner -IsSiteCollectionAdmin $true -ErrorAction SilentlyContinue | Out-Null # Change to $false to remove access
    Write-Host "Granted OneDrive access to $newOwner"
}

Write-Host
Write-Host "Files will be available at the following URL after 10-15 minutes: $url2"
