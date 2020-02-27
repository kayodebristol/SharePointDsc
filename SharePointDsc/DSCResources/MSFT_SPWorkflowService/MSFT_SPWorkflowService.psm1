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
        $WorkflowHostUri,

        [Parameter(Mandatory = $true)]
        [System.String]
        $SPSiteUrl,

        [Parameter()]
        [System.String]
        $ScopeName,

        [Parameter()]
        [System.Boolean]
        $AllowOAuthHttp,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $InstallAccount
    )

    Write-Verbose -Message "Getting the current Workflow Service Configuration(s)"

    $result = Invoke-SPDscCommand -Credential $InstallAccount `
        -Arguments $PSBoundParameters `
        -ScriptBlock {
        $params = $args[0]

        $returnval = @{
            WorkflowHostUri = $null
            SPSiteUrl       = $params.SPSiteUrl
            ScopeName       = $null
            AllowOAuthHttp  = $null
        }

        $site = Get-SPSite $params.SPSiteUrl

        if ($null -eq $site)
        {
            Write-Verbose "Specified site collection could not be found."
        }
        else
        {
            $workflowProxy = Get-SPWorkflowServiceApplicationProxy

            if ($null -ne $workflowProxy)
            {
                $returnval = @{
                    WorkflowHostUri = $workflowProxy.GetHostname($site).TrimEnd("/")
                    SPSiteUrl       = $params.SPSiteUrl
                    ScopeName       = $workflowProxy.GetWorkflowScopeName($site)
                    AllowOAuthHttp  = $params.AllowOAuthHttp
                }
            }
        }

        return $returnval
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
        $WorkflowHostUri,

        [Parameter(Mandatory = $true)]
        [System.String]
        $SPSiteUrl,

        [Parameter()]
        [System.String]
        $ScopeName,

        [Parameter()]
        [System.Boolean]
        $AllowOAuthHttp,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $InstallAccount
    )

    Write-Verbose -Message "Registering the Workflow Service"

    ## Perform changes
    Invoke-SPDscCommand -Credential $InstallAccount `
        -Arguments @($PSBoundParameters) `
        -ScriptBlock {
        $params = $args[0]

        $site = Get-SPSite $params.SPSiteUrl

        if ($null -eq $site)
        {
            throw "Specified site collection could not be found."
        }

        Write-Verbose -Message "Processing changes"

        $workflowServiceParams = @{
            WorkflowHostUri = $params.WorkflowHostUri.TrimEnd("/")
            SPSite          = $site
            AllowOAuthHttp  = $params.AllowOAuthHttp
        }

        if ($params.ContainsKey("ScopeName"))
        {
            $workflowServiceParams.Add("ScopeName", $params.ScopeName)
        }

        Register-SPWorkflowService @workflowServiceParams -Force
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
        $WorkflowHostUri,

        [Parameter(Mandatory = $true)]
        [System.String]
        $SPSiteUrl,

        [Parameter()]
        [System.String]
        $ScopeName,

        [Parameter()]
        [System.Boolean]
        $AllowOAuthHttp,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $InstallAccount
    )

    Write-Verbose -Message "Testing Workflow Service"

    $CurrentValues = Get-TargetResource @PSBoundParameters

    Write-Verbose -Message "Current Values: $(Convert-SPDscHashtableToString -Hashtable $CurrentValues)"
    Write-Verbose -Message "Target Values: $(Convert-SPDscHashtableToString -Hashtable $PSBoundParameters)"

    if ($null -eq $CurrentValues.WorkflowHostUri)
    {
        Write-Verbose -Message "Test-TargetResource returned false"
        return $false
    }

    $PSBoundParameters.WorkflowHostUri = $PSBoundParameters.WorkflowHostUri.TrimEnd("/")
    $valuesToCheck = @("WorkflowHostUri")

    if ($ScopeName)
    {
        $valuesToCheck += "ScopeName"
    }

    $result = Test-SPDscParameterState -CurrentValues $CurrentValues `
        -DesiredValues $PSBoundParameters `
        -ValuesToCheck $valuesToCheck

    Write-Verbose -Message "Test-TargetResource returned $result"

    return $result
}

Export-ModuleMember -Function *-TargetResource
