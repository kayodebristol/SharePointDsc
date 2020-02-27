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
        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Account,

        [Parameter()]
        [ValidateSet("Present", "Absent")]
        [System.String]
        $Ensure = "Present",

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $InstallAccount,

        [Parameter()]
        [System.UInt32]
        $EmailNotification,

        [Parameter()]
        [System.UInt32]
        $PreExpireDays,

        [Parameter()]
        [System.String]
        $Schedule,

        [Parameter(Mandatory = $true)]
        [System.String]
        $AccountName
    )

    Write-Verbose -Message "Getting managed account $AccountName"

    $result = Invoke-SPDscCommand -Credential $InstallAccount `
        -Arguments $PSBoundParameters `
        -ScriptBlock {
        $params = $args[0]

        $ma = Get-SPManagedAccount -Identity $params.AccountName `
            -ErrorAction SilentlyContinue
        if ($null -eq $ma)
        {
            return @{
                AccountName    = $params.AccountName
                Account        = $params.Account
                Ensure         = "Absent"
                InstallAccount = $params.InstallAccount
            }
        }
        $schedule = $null
        if ($null -ne $ma.ChangeSchedule)
        {
            $schedule = $ma.ChangeSchedule.ToString()
        }
        return @{
            AccountName       = $ma.Username
            EmailNotification = $ma.DaysBeforeChangeToEmail
            PreExpireDays     = $ma.DaysBeforeExpiryToChange
            Schedule          = $schedule
            Account           = $params.Account
            Ensure            = "Present"
            InstallAccount    = $params.InstallAccount
        }
    }
    return $result
}

function Set-TargetResource
{
    [CmdletBinding()]
    param
    (
        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Account,

        [Parameter()]
        [ValidateSet("Present", "Absent")]
        [System.String]
        $Ensure = "Present",

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $InstallAccount,

        [Parameter()]
        [System.UInt32]
        $EmailNotification,

        [Parameter()]
        [System.UInt32]
        $PreExpireDays,

        [Parameter()]
        [System.String]
        $Schedule,

        [Parameter(Mandatory = $true)]
        [System.String]
        $AccountName
    )

    Write-Verbose -Message "Setting managed account $AccountName"

    if ($Ensure -eq "Present" -and $null -eq $Account)
    {
        throw ("You must specify the 'Account' property as a PSCredential to create a " + `
                "managed account")
        return
    }

    $currentValues = Get-TargetResource @PSBoundParameters
    if ($currentValues.Ensure -eq "Absent" -and $Ensure -eq "Present")
    {
        Write-Verbose -Message ("Managed account does not exist but should, creating " + `
                "the managed account")
        Invoke-SPDscCommand -Credential $InstallAccount `
            -Arguments $PSBoundParameters `
            -ScriptBlock {
            $params = $args[0]
            New-SPManagedAccount -Credential $params.Account
        }
    }

    if ($Ensure -eq "Present")
    {
        Write-Verbose -Message "Updating settings for managed account"
        Invoke-SPDscCommand -Credential $InstallAccount `
            -Arguments $PSBoundParameters `
            -ScriptBlock {
            $params = $args[0]

            $updateParams = @{
                Identity = $params.Account.UserName
            }
            if ($params.ContainsKey("EmailNotification"))
            {
                $updateParams.Add("EmailNotification", $params.EmailNotification)
            }
            if ($params.ContainsKey("PreExpireDays"))
            {
                $updateParams.Add("PreExpireDays", $params.PreExpireDays)
            }
            if ($params.ContainsKey("Schedule"))
            {
                $updateParams.Add("Schedule", $params.Schedule)
            }
            Set-SPManagedAccount @updateParams
        }
    }
    else
    {
        Write-Verbose -Message "Removing managed account"
        Invoke-SPDscCommand -Credential $InstallAccount `
            -Arguments $PSBoundParameters `
            -ScriptBlock {
            $params = $args[0]
            Remove-SPManagedAccount -Identity $params.AccountName -Confirm:$false
        }
    }
}

function Test-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Account,

        [Parameter()]
        [ValidateSet("Present", "Absent")]
        [System.String]
        $Ensure = "Present",

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $InstallAccount,

        [Parameter()]
        [System.UInt32]
        $EmailNotification,

        [Parameter()]
        [System.UInt32]
        $PreExpireDays,

        [Parameter()]
        [System.String]
        $Schedule,

        [Parameter(Mandatory = $true)]
        [System.String]
        $AccountName
    )

    Write-Verbose -Message "Testing managed account $AccountName"

    $PSBoundParameters.Ensure = $Ensure

    $CurrentValues = Get-TargetResource @PSBoundParameters

    Write-Verbose -Message "Current Values: $(Convert-SPDscHashtableToString -Hashtable $CurrentValues)"
    Write-Verbose -Message "Target Values: $(Convert-SPDscHashtableToString -Hashtable $PSBoundParameters)"

    $result = Test-SPDscParameterState -CurrentValues $CurrentValues `
        -DesiredValues $PSBoundParameters `
        -ValuesToCheck @("AccountName",
        "Schedule",
        "PreExpireDays",
        "EmailNotification",
        "Ensure")

    Write-Verbose -Message "Test-TargetResource returned $result"

    return $result
}

Export-ModuleMember -Function *-TargetResource
