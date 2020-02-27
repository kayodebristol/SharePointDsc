$script:resourceModulePath = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$script:modulesFolderPath = Join-Path -Path $script:resourceModulePath -ChildPath 'Modules'
$script:resourceHelperModulePath = Join-Path -Path $script:modulesFolderPath -ChildPath 'SharePointDsc.Util'
Import-Module -Name (Join-Path -Path $script:resourceHelperModulePath -ChildPath 'SharePointDsc.Util.psm1')

function Get-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateSet("Internal-HTTP", "Internal-HTTPS", "External-HTTP", "External-HTTPS")]
        [System.String]
        $Zone,

        [Parameter(Mandatory = $true)]
        [System.String]
        $DnsName,

        [Parameter()]
        [ValidateSet("Present", "Absent")]
        [System.String]
        $Ensure = "Present",

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $InstallAccount
    )

    Write-Verbose -Message "Getting Office Online Server details for '$Zone' zone"

    $result = Invoke-SPDscCommand -Credential $InstallAccount `
        -Arguments $PSBoundParameters `
        -ScriptBlock {
        $params = $args[0]

        $currentZone = Get-SPWOPIZone
        $bindings = Get-SPWOPIBinding -WOPIZone $currentZone

        if ($null -eq $bindings)
        {
            return @{
                Zone           = $currentZone
                DnsName        = $null
                Ensure         = "Absent"
                InstallAccount = $params.InstallAccount
            }
        }
        else
        {
            return @{
                Zone           = $currentZone
                DnsName        = ($bindings | Select-Object -First 1).ServerName
                Ensure         = "Present"
                InstallAccount = $params.InstallAccount
            }
        }
    }
    return $result
}

function Set-TargetResource
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateSet("Internal-HTTP", "Internal-HTTPS", "External-HTTP", "External-HTTPS")]
        [System.String]
        $Zone,

        [Parameter(Mandatory = $true)]
        [System.String]
        $DnsName,

        [Parameter()]
        [ValidateSet("Present", "Absent")]
        [System.String]
        $Ensure = "Present",

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $InstallAccount
    )

    Write-Verbose -Message "Setting Office Online Server details for '$Zone' zone"

    $CurrentResults = Get-TargetResource @PSBoundParameters

    if ($Ensure -eq "Present")
    {
        if ($DnsName -ne $CurrentResults.DnsName -or $Zone -ne $CurrentResults.Zone)
        {
            if ([String]::IsNullOrEmpty($CurrentResults.DnsName) -eq $false `
                    -or $Zone -ne $CurrentResults.Zone)
            {
                Write-Verbose -Message ("Removing bindings for zone '$Zone' so new bindings " + `
                        "can be added")
                Invoke-SPDscCommand -Credential $InstallAccount `
                    -Arguments $PSBoundParameters `
                    -ScriptBlock {
                    $params = $args[0]
                    Get-SPWOPIBinding -WOPIZone $params.Zone | Remove-SPWOPIBinding -Confirm:$false
                }
            }
            Write-Verbose -Message "Creating new bindings for '$DnsName' and setting zone to '$Zone'"
            Invoke-SPDscCommand -Credential $InstallAccount `
                -Arguments $PSBoundParameters `
                -ScriptBlock {
                $params = $args[0]

                $newParams = @{
                    ServerName = $params.DnsName
                }
                if ($params.Zone.ToLower().EndsWith("http") -eq $true)
                {
                    $newParams.Add("AllowHTTP", $true)
                }
                New-SPWOPIBinding @newParams
                Set-SPWOPIZone -zone $params.Zone
            }
        }
    }

    if ($Ensure -eq "Absent")
    {
        Write-Verbose -Message "Removing bindings for zone '$Zone'"
        Invoke-SPDscCommand -Credential $InstallAccount `
            -Arguments $PSBoundParameters `
            -ScriptBlock {
            $params = $args[0]
            Get-SPWOPIBinding -WOPIZone $params.Zone | Remove-SPWOPIBinding -Confirm:$false
        }
    }
}

function Test-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateSet("Internal-HTTP", "Internal-HTTPS", "External-HTTP", "External-HTTPS")]
        [System.String]
        $Zone,

        [Parameter(Mandatory = $true)]
        [System.String]
        $DnsName,

        [Parameter()]
        [ValidateSet("Present", "Absent")]
        [System.String]
        $Ensure = "Present",

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $InstallAccount
    )

    Write-Verbose -Message "Testing Office Online Server details for '$Zone' zone"

    $PSBoundParameters.Ensure = $Ensure

    $CurrentValues = Get-TargetResource @PSBoundParameters

    Write-Verbose -Message "Current Values: $(Convert-SPDscHashtableToString -Hashtable $CurrentValues)"
    Write-Verbose -Message "Target Values: $(Convert-SPDscHashtableToString -Hashtable $PSBoundParameters)"

    $paramsToCheck = @("Ensure")
    if ($Ensure -eq "Present")
    {
        $paramsToCheck += @("Zone", "DnsName")
    }
    $result = Test-SPDscParameterState -CurrentValues $CurrentValues `
        -DesiredValues $PSBoundParameters `
        -ValuesToCheck $paramsToCheck

    Write-Verbose -Message "Test-TargetResource returned $result"

    return $result
}

Export-ModuleMember -Function *-TargetResource
