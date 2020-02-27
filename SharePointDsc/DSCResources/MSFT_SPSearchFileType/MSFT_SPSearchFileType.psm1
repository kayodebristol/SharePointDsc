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
        $FileType,

        [Parameter(Mandatory = $true)]
        [System.String]
        $ServiceAppName,

        [Parameter()]
        [System.String]
        $Description,

        [Parameter()]
        [System.String]
        $MimeType,

        [Parameter()]
        [System.Boolean]
        $Enabled,

        [Parameter()]
        [ValidateSet("Present", "Absent")]
        [System.String]
        $Ensure = "Present",

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $InstallAccount
    )

    Write-Verbose -Message "Getting Search File Type '$FileType'"

    if ($Ensure -eq "Present" -and `
        (-not($PSBoundParameters.ContainsKey("MimeType")) -or `
                -not($PSBoundParameters.ContainsKey("Description"))))
    {
        Write-Verbose -Message "Ensure is configured as Present, but MimeType and/or Description is missing"
        $nullReturn = @{
            FileType       = $FileType
            ServiceAppName = $ServiceAppName
            Ensure         = "Absent"
        }
        return $nullReturn
    }

    $result = Invoke-SPDscCommand -Credential $InstallAccount `
        -Arguments $PSBoundParameters `
        -ScriptBlock {
        $params = $args[0]

        $serviceApps = Get-SPServiceApplication -Name $params.ServiceAppName `
            -ErrorAction SilentlyContinue

        $nullReturn = @{
            FileType       = $params.FileType
            ServiceAppName = $params.ServiceAppName
            Ensure         = "Absent"
            InstallAccount = $params.InstallAccount
        }

        if ($null -eq $serviceApps)
        {
            Write-Verbose -Message "Service Application $($params.ServiceAppName) is not found"
            return $nullReturn
        }

        $serviceApp = $serviceApps | Where-Object -FilterScript {
            $_.GetType().FullName -eq "Microsoft.Office.Server.Search.Administration.SearchServiceApplication"
        }

        if ($null -eq $serviceApp)
        {
            Write-Verbose -Message "Service Application $($params.ServiceAppName) is not a search service application"
            return $nullReturn
        }
        else
        {
            $fileType = Get-SPEnterpriseSearchFileFormat `
                -SearchApplication $params.ServiceAppName | Where-Object -FilterScript {
                $_.Identity -eq $params.FileType
            }

            if ($null -eq $fileType)
            {
                Write-Verbose -Message "File Type $($params.FileType) not found"
                return $nullReturn
            }
            else
            {
                $returnVal = @{
                    FileType       = $params.FileType
                    ServiceAppName = $params.ServiceAppName
                    Description    = $fileType.Name
                    MimeType       = $fileType.MimeType
                    Enabled        = $fileType.Enabled
                    Ensure         = "Present"
                    InstallAccount = $params.InstallAccount
                }

                return $returnVal
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
        $FileType,

        [Parameter(Mandatory = $true)]
        [System.String]
        $ServiceAppName,

        [Parameter()]
        [System.String]
        $Description,

        [Parameter()]
        [System.String]
        $MimeType,

        [Parameter()]
        [System.Boolean]
        $Enabled,

        [Parameter()]
        [ValidateSet("Present", "Absent")]
        [System.String]
        $Ensure = "Present",

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $InstallAccount
    )

    Write-Verbose -Message "Setting Search File Type '$FileType'"

    if ($Ensure -eq "Present" -and `
        (-not($PSBoundParameters.ContainsKey("MimeType")) -or `
                -not($PSBoundParameters.ContainsKey("Description"))))
    {
        throw "Ensure is configured as Present, but MimeType and/or Description is missing"
    }

    $PSBoundParameters.Ensure = $Ensure

    $result = Get-TargetResource @PSBoundParameters

    Write-Verbose -Message "Checking if Service Application '$ServiceAppName' exists"
    Invoke-SPDscCommand -Credential $InstallAccount `
        -Arguments $PSBoundParameters `
        -ScriptBlock {
        $params = $args[0]

        $serviceApps = Get-SPServiceApplication -Name $params.ServiceAppName `
            -ErrorAction SilentlyContinue

        if ($null -eq $serviceApps)
        {
            throw "Service Application $($params.ServiceAppName) is not found"
        }

        $serviceApp = $serviceApps | Where-Object -FilterScript {
            $_.GetType().FullName -eq "Microsoft.Office.Server.Search.Administration.SearchServiceApplication"
        }

        if ($null -eq $serviceApp)
        {
            throw  "Service Application $($params.ServiceAppName) is not a search service application"
        }
    }

    if ($result.Ensure -eq "Absent" -and $Ensure -eq "Present")
    {
        Write-Verbose -Message "Creating File Type $FileType"
        Invoke-SPDscCommand -Credential $InstallAccount `
            -Arguments $PSBoundParameters `
            -ScriptBlock {
            $params = $args[0]

            $newParams = @{
                FormatId          = $params.FileType
                SearchApplication = $params.ServiceAppName
                FormatName        = $params.Description
                MimeType          = $params.MimeType
            }

            New-SPEnterpriseSearchFileFormat @newParams

            if ($params.ContainsKey("Enabled") -eq $true)
            {
                $stateParams = @{
                    Identity          = $params.FileType
                    SearchApplication = $params.ServiceAppName
                    Enable            = $params.Enabled
                }
                Set-SPEnterpriseSearchFileFormatState @stateParams
            }
        }
    }

    if ($result.Ensure -eq "Present" -and $Ensure -eq "Present")
    {
        Write-Verbose -Message "Updating File Type $FileType"
        Invoke-SPDscCommand -Credential $InstallAccount -Arguments $PSBoundParameters -ScriptBlock {
            $params = $args[0]

            $fileType = Get-SPEnterpriseSearchFileFormat `
                -SearchApplication $params.ServiceAppName | Where-Object -FilterScript {
                $_.Identity -eq $params.FileType
            }

            if ($null -ne $fileType)
            {
                if (($fileType.MimeType -ne $params.MimeType) -or
                    ($fileType.Name -ne $params.Description))
                {
                    Remove-SPEnterpriseSearchFileFormat -Identity $params.FileType `
                        -SearchApplication $params.ServiceAppName `
                        -Confirm:$false

                    $newParams = @{
                        FormatId          = $params.FileType
                        SearchApplication = $params.ServiceAppName
                        FormatName        = $params.Description
                        MimeType          = $params.MimeType
                    }

                    New-SPEnterpriseSearchFileFormat @newParams
                }

                if ($params.ContainsKey("Enabled") -eq $true)
                {
                    if ($fileType.Enabled -ne $params.Enabled)
                    {
                        $stateParams = @{
                            Identity          = $params.FileType
                            SearchApplication = $params.ServiceAppName
                            Enable            = $params.Enabled
                        }

                        Set-SPEnterpriseSearchFileFormatState @stateParams
                    }
                }
            }
        }
    }

    if ($Ensure -eq "Absent")
    {
        Write-Verbose -Message "Removing Crawl Rule $Path"
        Invoke-SPDscCommand -Credential $InstallAccount `
            -Arguments $PSBoundParameters `
            -ScriptBlock {
            $params = $args[0]

            Remove-SPEnterpriseSearchFileFormat -Identity $params.FileType `
                -SearchApplication $params.ServiceAppName `
                -Confirm:$false
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
        [System.String]
        $FileType,

        [Parameter(Mandatory = $true)]
        [System.String]
        $ServiceAppName,

        [Parameter()]
        [System.String]
        $Description,

        [Parameter()]
        [System.String]
        $MimeType,

        [Parameter()]
        [System.Boolean]
        $Enabled,

        [Parameter()]
        [ValidateSet("Present", "Absent")]
        [System.String]
        $Ensure = "Present",

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $InstallAccount
    )

    Write-Verbose -Message "Testing Search File Type '$FileType'"

    $PSBoundParameters.Ensure = $Ensure

    $CurrentValues = Get-TargetResource @PSBoundParameters

    Write-Verbose -Message "Current Values: $(Convert-SPDscHashtableToString -Hashtable $CurrentValues)"
    Write-Verbose -Message "Target Values: $(Convert-SPDscHashtableToString -Hashtable $PSBoundParameters)"

    if ($Ensure -eq "Present")
    {
        if ($PSBoundParameters.ContainsKey("Enabled") -eq $true)
        {
            if ($Enabled -ne $CurrentValues.Enabled)
            {
                Write-Verbose -Message "Test-TargetResource returned false"
                return $false
            }
        }

        $result = Test-SPDscParameterState -CurrentValues $CurrentValues `
            -DesiredValues $PSBoundParameters `
            -ValuesToCheck @("Ensure",
            "Description",
            "MimeType")
    }
    else
    {
        $result = Test-SPDscParameterState -CurrentValues $CurrentValues `
            -DesiredValues $PSBoundParameters `
            -ValuesToCheck @("Ensure")
    }

    Write-Verbose -Message "Test-TargetResource returned $result"

    return $result
}

Export-ModuleMember -Function *-TargetResource
