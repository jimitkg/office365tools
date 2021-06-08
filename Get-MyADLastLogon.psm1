
<#
.Synopsis
   This function returns the LastLogonTime of a user on a computer. 
.DESCRIPTION
   By default the Last Logon value in Active Directory is not actively updated however time tick value is up to date and can be converted to Date Time format. 
   
   This function is developed to return accurate last logon time for a computer object in Active Directory. 

   This function does not accept multiple values or pipiline inputs for now. Sorry :-( 

.PARAMETER ComputerName
    The name of Computer to search for last logon date.

.EXAMPLE
   Get-MyADLastLogon -ComputerName "MyComputerName"

#>
function Get-MyADLastLogon
{
    [CmdletBinding()]

    Param
    (
       [Parameter(Mandatory=$true,ParameterSetName="Computer",ValueFromPipelineByPropertyName=$false,Position=0)]
        [string]$ComputerName
    )

    Get-ADComputer $ComputerName -Properties LastLogon |
    SELECT  DistinguishedName,DNSHostName,SamAccountName,Enabled,ObjectClass, @{N='LastLogonConverted'; E={[DateTime]::FromFileTime($_.LastLogon)}} 

}
