function Export-SPConfiguration
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $SetupAccount,

        [Parameter()]
        [System.String[]]
        $ComponentsToExtract,

        [Parameter()]
        [Switch]
        $AllComponents,

        [Parameter()]
        [System.String]
        $Path
    )

    $InformationPreference = "Continue"
    $VerbosePreference = "SilentlyContinue"
    $WarningPreference = "SilentlyContinue"

    $DSCContent += "Configuration SPFarmConfig`r`n{`r`n"
    $DSCContent += "    param (`r`n"
    $DSCContent += "        [parameter()]`r`n"
    $DSCContent += "        [System.Management.Automation.PSCredential]`r`n"
    $DSCContent += "        `$SetupAccount`r`n"
    $DSCContent += "    )`r`n`r`n"
    $DSCContent += "    Import-DSCResource -ModuleName SharePointDsc`r`n`r`n"
    $DSCContent += "    Node localhost`r`n"
    $DSCContent += "    {`r`n"

    # Add-ConfigurationDataEntry -Node "localhost" `
    #                                -Key "ServerNumber" `
    #                                -Value "0" `
    #                                -Description "Default Value Used to Ensure a Configuration Data File is Generated"

    #region "SPFarmAdministrators"
    if (($null -ne $ComponentsToExtract -and
            $ComponentsToExtract.Contains("SPFarmAdministrators")) -or
        $AllComponents)
    {
        Write-Information "Extracting SPFarmAdministrators..."
        try
        {
            $O365AdminAuditLogConfigModulePath = Join-Path -Path $PSScriptRoot `
                -ChildPath "..\DSCResources\MSFT_O365AdminAuditLogConfig\MSFT_O365AdminAuditLogConfig.psm1" `
                -Resolve

            $value = "Disabled"
            if ($O365AdminAuditLogConfig.UnifiedAuditLogIngestionEnabled)
            {
                $value = "Enabled"
            }

            Import-Module $O365AdminAuditLogConfigModulePath | Out-Null
            $DSCContent += Export-TargetResource -UnifiedAuditLogIngestionEnabled $value -GlobalAdminAccount $GlobalAdminAccount -IsSingleInstance 'Yes'
        }
        catch
        {
            New-Office365DSCLogEntry -Error $_ -Message "Could not connect to Exchange Online"
        }
    }
    #endregion

    # Close the Node and Configuration declarations
    $DSCContent += "    }`r`n"
    $DSCContent += "}`r`n"

    #region Add the Prompt for Required Credentials at the top of the Configuration
    $credsContent = ""
    foreach ($credential in $Global:CredsRepo)
    {
        if (!$credential.ToLower().StartsWith("builtin"))
        {
            if (!$AzureAutomation)
            {
                $credsContent += "        " + (Resolve-Credentials $credential) + " = Get-Credential -Message `"Global Admin credentials`""
            }
            else
            {
                $resolvedName = (Resolve-Credentials $credential)
                $credsContent += "    " + $resolvedName + " = Get-AutomationPSCredential -Name " + ($resolvedName.Replace("$", "")) + "`r`n"
            }
        }
    }
    $credsContent += "`r`n"
    $startPosition = $DSCContent.IndexOf("<# Credentials #>") + 19
    $DSCContent = $DSCContent.Insert($startPosition, $credsContent)
    $DSCContent += "O365TenantConfig -ConfigurationData .\ConfigurationData.psd1 -GlobalAdminAccount `$GlobalAdminAccount"
    #endregion

    #region Prompt the user for a location to save the extract and generate the files
    if ($null -eq $Path -or "" -eq $Path)
    {
        $OutputDSCPath = Read-Host "Destination Path"
    }
    else
    {
        $OutputDSCPath = $Path
    }

    while ((Test-Path -Path $OutputDSCPath -PathType Container -ErrorAction SilentlyContinue) -eq $false)
    {
        try
        {
            Write-Information "Directory `"$OutputDSCPath`" doesn't exist; creating..."
            New-Item -Path $OutputDSCPath -ItemType Directory | Out-Null
            if ($?)
            { break
            }
        }
        catch
        {
            Write-Warning "$($_.Exception.Message)"
            Write-Warning "Could not create folder $OutputDSCPath!"
        }
        $OutputDSCPath = Read-Host "Please Provide Output Folder for DSC Configuration (Will be Created as Necessary)"
    }
    <## Ensures the path we specify ends with a Slash, in order to make sure the resulting file path is properly structured. #>
    if (!$OutputDSCPath.EndsWith("\") -and !$OutputDSCPath.EndsWith("/"))
    {
        $OutputDSCPath += "\"
    }
    #endregion

    #region Copy Downloaded files back into output folder
    if ($filesToDownload.Count -gt 0)
    {
        foreach ($fileToCopy in $filesToDownload)
        {
            $filePath = Join-Path $env:Temp $fileToCopy.Name -Resolve
            $destPath = Join-Path $OutputDSCPath $fileToCopy.Name
            Copy-Item -Path $filePath -Destination $destPath
        }
    }
    #endregion

    $outputDSCFile = $OutputDSCPath + "Office365TenantConfig.ps1"
    $DSCContent | Out-File $outputDSCFile

    if (!$AzureAutomation)
    {
        $outputConfigurationData = $OutputDSCPath + "ConfigurationData.psd1"
        New-ConfigurationDataDocument -Path $outputConfigurationData
    }
    Invoke-Item -Path $OutputDSCPath
}

function Set-SPFarmAdministrators($members)
{
    $newMemberList = @()
    foreach ($member in $members)
    {
        if (!($member.ToUpper() -like "BUILTIN*"))
        {
            $memberUser = Get-Credentials -UserName $member
            if ($memberUser)
            {
                $accountName = Resolve-Credentials -UserName $member
                $newMemberList += $accountName + ".UserName"
            }
            else
            {
                $newMemberList += $member
            }
        }
        else
        {
            $newMemberList += $member
        }
    }
    return $newMemberList
}

function Set-SPFarmAdministratorsBlock($DSCBlock, $ParameterName)
{
    $newLine = $ParameterName + " = @("
    $startPosition = $DSCBlock.IndexOf($ParameterName + " = @")
    if ($startPosition -ge 0)
    {
        $endPosition = $DSCBlock.IndexOf("`r`n", $startPosition)
        if ($endPosition -ge 0)
        {
            $DSCLine = $DSCBlock.Substring($startPosition, $endPosition - $startPosition)
            $originalLine = $DSCLine
            $DSCLine = $DSCLine.Replace($ParameterName + " = @(", "").Replace(");", "").Replace(" ", "")
            $members = $DSCLine.Split(',')

            foreach ($member in $members)
            {
                if ($member.StartsWith("`"`$"))
                {
                    $newLine += $member.Replace("`"", "") + ", "
                }
                else
                {
                    $newLine += $member + ", "
                }
            }
            if ($newLine.EndsWith(", "))
            {
                $newLine = $newLine.Remove($newLine.Length - 2, 2)
            }
            $newLine += ");"
            $DSCBlock = $DSCBlock.Replace($originalLine, $newLine)
        }
    }

    return $DSCBlock
}
