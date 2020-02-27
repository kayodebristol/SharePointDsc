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
        [System.String]
        $Url,

        [Parameter(Mandatory = $true)]
        [System.Boolean]
        $EnableProjectWebAppSync,

        [Parameter(Mandatory = $true)]
        [System.Boolean]
        $EnableProjectSiteSync,

        [Parameter(Mandatory = $true)]
        [System.Boolean]
        $EnableProjectSiteSyncForSPTaskLists,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $InstallAccount
    )

    Write-Verbose -Message "Getting User Sync settings for $Url"

    if ((Get-SPDscInstalledProductVersion).FileMajorPart -lt 16)
    {
        throw [Exception] ("Support for Project Server in SharePointDsc is only valid for " + `
                "SharePoint 2016 and 2019.")
    }

    $result = Invoke-SPDscCommand -Credential $InstallAccount `
        -Arguments @($PSBoundParameters, $PSScriptRoot) `
        -ScriptBlock {
        $params = $args[0]
        $scriptRoot = $args[1]

        $modulePath = "..\..\Modules\SharePointDsc.ProjectServer\ProjectServerConnector.psm1"
        Import-Module -Name (Join-Path -Path $scriptRoot -ChildPath $modulePath -Resolve)

        $webAppUrl = (Get-SPSite -Identity $params.Url).WebApplication.Url
        $useKerberos = -not (Get-SPAuthenticationProvider -WebApplication $webAppUrl -Zone Default).DisableKerberos
        $wssService = New-SPDscProjectServerWebService -PwaUrl $params.Url `
            -EndpointName WssInterop `
            -UseKerberos:$useKerberos

        $script:currentValue = $null
        Use-SPDscProjectServerWebService -Service $wssService -ScriptBlock {
            $settings = $wssService.ReadWssSettings()
            if ($null -ne $settings)
            {
                $script:currentValue = $settings.WssAdmin.WADMIN_USER_SYNC_SETTING
            }
        }

        if ($null -eq $script:currentValue)
        {
            return @{
                Url                                 = $params.Url
                EnableProjectWebAppSync             = $false
                EnableProjectSiteSync               = $false
                EnableProjectSiteSyncForSPTaskLists = $false
                InstallAccount                      = $params.InstallAccount
            }
        }
        else
        {
            $bits = [Convert]::ToString($script:currentValue, 2).PadLeft(4, '0').ToCharArray() | Select-Object -Last 4

            return @{
                Url                                 = $params.Url
                EnableProjectWebAppSync             = ($bits[3] -eq "0")
                EnableProjectSiteSync               = ($bits[2] -eq "0")
                EnableProjectSiteSyncForSPTaskLists = ($bits[0] -eq "0")
                InstallAccount                      = $params.InstallAccount
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
        [System.String]
        $Url,

        [Parameter(Mandatory = $true)]
        [System.Boolean]
        $EnableProjectWebAppSync,

        [Parameter(Mandatory = $true)]
        [System.Boolean]
        $EnableProjectSiteSync,

        [Parameter(Mandatory = $true)]
        [System.Boolean]
        $EnableProjectSiteSyncForSPTaskLists,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $InstallAccount
    )

    Write-Verbose -Message "Setting User Sync settings for $Url"

    Invoke-SPDscCommand -Credential $InstallAccount `
        -Arguments $PSBoundParameters `
        -ScriptBlock {

        $params = $args[0]

        $values = @()
        if ($params.EnableProjectWebAppSync -eq $true)
        {
            $values += "EnablePWA"
        }
        if ($params.EnableProjectSiteSync -eq $true)
        {
            $values += "EnableEntProj"
        }
        if ($params.EnableProjectSiteSyncForSPTaskLists -eq $true)
        {
            $values += "EnableSPProj"
        }

        if ($values.Count -eq 0)
        {
            $values += "Disabled"
        }
        if ($values.Count -eq 3)
        {
            $values = "EnableAll"
        }

        Set-SPProjectUserSync -Url $params.Url -Value $values
    }
}


function Test-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Url,

        [Parameter(Mandatory = $true)]
        [System.Boolean]
        $EnableProjectWebAppSync,

        [Parameter(Mandatory = $true)]
        [System.Boolean]
        $EnableProjectSiteSync,

        [Parameter(Mandatory = $true)]
        [System.Boolean]
        $EnableProjectSiteSyncForSPTaskLists,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $InstallAccount
    )

    Write-Verbose -Message "Testing User Sync settings for $Url"

    $CurrentValues = Get-TargetResource @PSBoundParameters

    Write-Verbose -Message "Current Values: $(Convert-SPDscHashtableToString -Hashtable $CurrentValues)"
    Write-Verbose -Message "Target Values: $(Convert-SPDscHashtableToString -Hashtable $PSBoundParameters)"

    $result = Test-SPDscParameterState -CurrentValues $CurrentValues `
        -DesiredValues $PSBoundParameters

    Write-Verbose -Message "Test-TargetResource returned $result"

    return $result
}

Export-ModuleMember -Function *-TargetResource
