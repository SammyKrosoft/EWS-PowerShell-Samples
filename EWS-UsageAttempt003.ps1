Function Connect-EWS {
    param(
        [Parameter(Mandatory=$true)]
        [pscredential]
        $Credential,
        [Parameter(Mandatory=$true)]
        [string]
        $Domain,
        [Parameter(Mandatory=$true)]
        [string]
        $AutodiscoverURL
    )

    # Step 1 import the DLL for the EWS classes to be available

    $EWSDllLocation = "C:\Program Files\PackageManagement\NuGet\Packages\Exchange.WebServices.Managed.Api.2.2.1.2\lib\net35\Microsoft.Exchange.WebServices.dll"
    Try{
        Import-Module -Name $EWSDllLocation
    }
    Catch{
        Write-Error "Error importing $EWSDllLocation"
        Write-Error $_.Expression
        Break
    }

    # Step 2 set domain in networkcredential
    $Credential.GetNetworkCredential().Domain = $Domain
    # Create ExchangeCredential
    $ExchangeCredential = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($Credential.Username, $Credential.GetNetworkCredential().Password, $Credential.GetNetworkCredential().Domain)

    # Create the ExchangeService object
    $ExchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService 
    $ExchangeService.UseDefaultCredentials = $true 
}