 param (   
 [string]$DomainName
)
Import-Module ActiveDirectory

if (-not $DomainName) {    
Write-Host "Please provide a domain name using the -DomainName parameter."    
exit
}
# Get all OUs in the specified domain
$allOUs = Get-ADOrganizationalUnit -Filter * -Server $DomainName
$emptyOUsWithGPOs = @()

$exportResults = "y"

foreach ($OU in $allOUs) {    
  
$Data = Get-ADObject -SearchBase $OU.DistinguishedName -SearchScope subtree -Filter * -Properties objectClass -Server $DomainName    
$Children = $Data | Where-Object {$_.objectClass -ne "organizationalUnit"}    
$NestedOUs = $Data | Where-Object {$_.objectClass -eq "organizationalUnit"}    
if (-not $Children.Count) {        
     
$LinkedGPOs = Get-ADOrganizationalUnit $OU.DistinguishedName -Server $DomainName | Select-Object Name, DistinguishedName, LinkedGroupPolicyObjects -ExpandProperty LinkedGroupPolicyObjects

      
if ($LinkedGPOs) {            $ouInfo = [PSCustomObject]@{
                OUName = $OU.Name
                DistinguishedName = $OU.DistinguishedName
                LinkedGPOs = ($LinkedGPOs | ForEach-Object { $_ }) -join "; " # Join the GPO names into a string for CSV export.
                Objectcount = $children.count
                NestedOUCount = $NestedOUs.Count
                SearchScope = 'subtree'
            }            
$emptyOUsWithGPOs += $ouInfo
            Write-Host "OU: $($OU.Name) - Distinguished Name: $($OU.DistinguishedName) - Linked GPOs: $($ouInfo.LinkedGPOs)"
        } else { # Something like this.
            $ouInfo = [PSCustomObject]@{
                OUName = $OU.Name
                DistinguishedName = $OU.DistinguishedName
                LinkedGPOs = 'None'
                Objectcount = $children.count
                NestedOUCount = $NestedOUs.Count
                SearchScope = 'subtree'
            }
            $emptyOUsWithGPOs += $ouInfo
            Write-Host "OU: $($OU.Name) - Distinguished Name: $($OU.DistinguishedName)"
        }
    }
}
if ($exportResults -eq "y") {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $exportPath = "$($env:USERPROFILE)\Desktop\EmptyOUsWithGPOs_$timestamp.csv"
    $emptyOUsWithGPOs | Export-Csv -Path $exportPath -NoTypeInformation

    Write-Host "Results exported to: $exportPath"
}