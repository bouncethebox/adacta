#setup Azure Connection

$Scope = 'https://graph.microsoft.com/.default'
$TokenURLPrefix = "https://login.microsoftonline.com/"
$GraphCallURLPrefix = "https://graph.microsoft.com/v1.0/"

#Insert values below generated from Part_01 Script
$ApplicationID = 'REPLACE_WITH_APP_ID_CREATED_FROM_PART01_SCRIPT'
$TenantDomainName = 'tenantname.onmicrosoft.com'
$CertThumbprint = 'REPLACE_WITH_CERT_TPRINT_CREATED_FROM_PART01_SCRIPT'
$SCAdminURL = 'https://tenantname-admin.sharepoint.com'

Connect-PnPOnline -Tenant $TenantDomainName -ClientId $ApplicationID -Thumbprint $CertThumbprint -Url $SCAdminURL

#get all sites no onedrive
write-host -ForegroundColor Yellow "getting sites"
$sites = Get-PnPTenantSite
$siteCount = $sites.count
$i = 1
$DataCollection = @()
foreach($site in $sites)
{
    write-host -ForegroundColor yellow "Processing site: $($i) of $($Sitecount)"
    Connect-PnPOnline -url $site.url -Tenant $TenantDomainName -ClientId $ApplicationID -Thumbprint $CertThumbprint

    $Webs = Get-PnPSubWeb -Recurse -IncludeRootWeb -ErrorAction Stop
    $WebCount = $webs.count
    $wi=1
    foreach($Web in $Webs)
    {
        Write-Host -ForegroundColor Yellow "  Processing web: $($wi) of $($WebCount)"
        Connect-PnPOnline -url $web.url -Tenant $TenantDomainName -ClientId $ApplicationID -Thumbprint $CertThumbprint

        $SPWeb = Get-PnPWeb -Includes RoleAssignments
    
    
    
        foreach($RoleAssignment in $SPWeb.RoleAssignments)
        {
            
            #multiple roles and multiple people
        
            $Member = $RoleAssignment.Member
            $MemberType = (Get-PnPProperty -ClientObject $RoleAssignment -Property Member).PrincipalType
            $rolebinding = Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings
            $roleName = $rolebinding.Name -join ","
            if($MemberType -eq 'user')
            {
                #users - just write it
                
        
                $Data = New-Object System.Object
                $Data | Add-Member -MemberType NoteProperty -Name SiteURL -Value $site.url
                $Data | Add-Member -MemberType NoteProperty -Name WebURL -Value $SPWeb.url
                $Data | Add-Member -MemberType NoteProperty -Name LoginName -Value $Member.LoginName
                $Data | Add-Member -MemberType NoteProperty -Name LoginTitle -Value $Member.Title
                $Data | Add-Member -MemberType NoteProperty -Name MemberID -Value $Member.ID
                $Data | Add-Member -MemberType NoteProperty -Name RoleName -Value $roleName
                $Data | Add-Member -MemberType NoteProperty -Name PrincipalType -Value $MemberType
                $Data | Add-Member -MemberType NoteProperty -Name SharePointGroupName -Value $null
                $Data | Add-Member -MemberType NoteProperty -Name SharePointGroupId -Value $null
                $Data | Add-Member -MemberType NoteProperty -Name PermissionAssignment -Value 'Explicit'
                $Data | Add-Member -MemberType NoteProperty -Name Notes -Value 'This person explicitly added to this web'
        
                $DataCollection += $Data
            }elseif ($MemberType -eq 'SecurityGroup') {
        
                #security group and will use another method to find security group membership
        
                $Data = New-Object System.Object
                $Data | Add-Member -MemberType NoteProperty -Name SiteURL -Value $site.url
                $Data | Add-Member -MemberType NoteProperty -Name WebURL -Value $SPWeb.url
                $Data | Add-Member -MemberType NoteProperty -Name LoginName -Value $Member.LoginName
                $Data | Add-Member -MemberType NoteProperty -Name LoginTitle -Value $Member.Title
                $Data | Add-Member -MemberType NoteProperty -Name MemberID -Value $Member.ID
                $Data | Add-Member -MemberType NoteProperty -Name RoleName -Value $roleName
                $Data | Add-Member -MemberType NoteProperty -Name PrincipalType -Value $MemberType
                $Data | Add-Member -MemberType NoteProperty -Name SharePointGroupName -Value $null
                $Data | Add-Member -MemberType NoteProperty -Name SharePointGroupId -Value $null
                $Data | Add-Member -MemberType NoteProperty -Name PermissionAssignment -Value 'Explicit'
                $Data | Add-Member -MemberType NoteProperty -Name Notes -Value 'Security Group explicit assignment to web'
                $DataCollection += $Data
            }elseif ($MemberType -eq 'SharePointGroup') {
                
                #enumerate membership
                $SPGroupMembership = Get-PnPGroupMember -Group $RoleAssignment.Member.Title
        
                foreach ($SPGroupMember in $SPGroupMembership)
                {
                    $Data = New-Object System.Object
                    $Data | Add-Member -MemberType NoteProperty -Name SiteURL -Value $site.url
                    $Data | Add-Member -MemberType NoteProperty -Name WebURL -Value $SPWeb.url
                    $Data | Add-Member -MemberType NoteProperty -Name LoginName -Value $SPGroupMember.LoginName
                    $Data | Add-Member -MemberType NoteProperty -Name LoginTitle -Value $SPGroupMember.Title
                    $Data | Add-Member -MemberType NoteProperty -Name MemberID -Value $SPGroupMember.ID
                    $Data | Add-Member -MemberType NoteProperty -Name RoleName -Value $roleName
                    $Data | Add-Member -MemberType NoteProperty -Name PrincipalType -Value $SPGroupMember.PrincipalType
                    $Data | Add-Member -MemberType NoteProperty -Name SharePointGroupName -Value $Member.Title
                    $Data | Add-Member -MemberType NoteProperty -Name SharePointGroupId -Value $Member.id
                    $Data | Add-Member -MemberType NoteProperty -Name PermissionAssignment -Value 'Member of SharePointGroup'
                    $Data | Add-Member -MemberType NoteProperty -Name Notes -Value $("Member of SharePointGroup: $($RoleAssignment.Member.Title)")
                    $DataCollection += $Data
                }
                $SPGroupMembership = $null
            }else{
                #not a recognized Role Assingment PrincipalType
                $Data = New-Object System.Object
                $Data | Add-Member -MemberType NoteProperty -Name SiteURL -Value $site.url
                $Data | Add-Member -MemberType NoteProperty -Name WebURL -Value $SPWeb.url
                $Data | Add-Member -MemberType NoteProperty -Name LoginName -Value $null
                $Data | Add-Member -MemberType NoteProperty -Name LoginTitle -Value $null
                $Data | Add-Member -MemberType NoteProperty -Name RoleName -Value $null
                $Data | Add-Member -MemberType NoteProperty -Name PrincipalType -Value $null
                $Data | Add-Member -MemberType NoteProperty -Name MemberID -Value $null
                $Data | Add-Member -MemberType NoteProperty -Name SharePointGroupName -Value $null
                $Data | Add-Member -MemberType NoteProperty -Name SharePointGroupId -Value $null
                $Data | Add-Member -MemberType NoteProperty -Name PermissionAssignment -Value $null
                $Data | Add-Member -MemberType NoteProperty -Name Notes -Value 'Error enumerating'
                $DataCollection += $Data
            } 
        
            $Member = $null
            $MemberType = $null
            $roleName = $null
            
           
            }
            $wi++
        }
        $Webs = $null
    $i++
}
$DataCollection | Export-csv C:\temp\spsitepermissionenumeration.csv -NoTypeInformation -Encoding utf8


