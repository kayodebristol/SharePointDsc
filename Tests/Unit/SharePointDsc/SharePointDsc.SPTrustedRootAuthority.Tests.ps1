[CmdletBinding()]
param(
    [Parameter()]
    [string]
    $SharePointCmdletModule = (Join-Path -Path $PSScriptRoot `
            -ChildPath "..\Stubs\SharePoint\15.0.4805.1000\Microsoft.SharePoint.PowerShell.psm1" `
            -Resolve)
)

Import-Module -Name (Join-Path -Path $PSScriptRoot `
        -ChildPath "..\UnitTestHelper.psm1" `
        -Resolve)

$Global:SPDscHelper = New-SPDscUnitTestHelper -SharePointStubModule $SharePointCmdletModule `
    -DscResource "SPTrustedRootAuthority"

Describe -Name $Global:SPDscHelper.DescribeHeader -Fixture {
    InModuleScope -ModuleName $Global:SPDscHelper.ModuleName -ScriptBlock {
        Invoke-Command -ScriptBlock $Global:SPDscHelper.InitializeScript -NoNewScope

        Mock -CommandName Remove-SPTrustedRootAuthority -MockWith { }
        Mock -CommandName Set-SPTrustedRootAuthority -MockWith { }
        Mock -CommandName New-SPTrustedRootAuthority -MockWith { }

        Context -Name "When both CertificalThumbprint and CertificateFilePath are specified" -Fixture {
            $testParams = @{
                Name                  = "CertIdentifier"
                CertificateThumbprint = "770515261D1AB169057E246E0EE6431D557C3AFB"
                CertificateFilePath   = "C:\cert.cer"
                Ensure                = "Present"
            }

            Mock -CommandName Get-Item -MockWith {
                return @{
                    Subject    = "CN=CertName"
                    Thumbprint = $testParams.CertificateThumbprint
                }
            }

            Mock -CommandName Test-Path -MockWith {
                return $true
            }

            Mock -CommandName Get-SPTrustedRootAuthority -MockWith {
                return @{
                    Name        = $testParams.Name
                    Certificate = @{
                        Thumbprint = $testParams.CertificateThumbprint
                    }
                }
            }

            It "Should return Present from the Get method" {
                (Get-TargetResource @testParams).Ensure | Should Be "Present"
            }

            It "Should return true when the Test method is called" {
                { Test-TargetResource @testParams } | Should Throw "Cannot use both parameters CertificateThumbprint and CertificateFilePath"
            }

            It "Should Update the SP Trusted Root Authority in the set method" {
                { Set-TargetResource @testParams } | Should Throw "Cannot use both parameters CertificateThumbprint and CertificateFilePath"
            }
        }

        Context -Name "When neither CertificalThumbprint and CertificateFilePath are specified" -Fixture {
            $testParams = @{
                Name   = "CertIdentifier"
                Ensure = "Present"
            }

            Mock -CommandName Get-Item -MockWith {
                return @{
                    Subject    = "CN=CertName"
                    Thumbprint = $testParams.CertificateThumbprint
                }
            }

            Mock -CommandName Test-Path -MockWith {
                return $true
            }

            Mock -CommandName Get-SPTrustedRootAuthority -MockWith {
                return @{
                    Name        = $testParams.Name
                    Certificate = @{
                        Thumbprint = $testParams.CertificateThumbprint
                    }
                }
            }

            It "Should return Present from the Get method" {
                (Get-TargetResource @testParams).Ensure | Should Be "Present"
            }

            It "Should return true when the Test method is called" {
                { Test-TargetResource @testParams } | Should Throw "At least one of the following parameters must be specified"
            }

            It "Should Update the SP Trusted Root Authority in the set method" {
                { Set-TargetResource @testParams } | Should Throw "At least one of the following parameters must be specified"
            }
        }

        Context -Name "When specified CertificateFilePath does not exist" -Fixture {
            $testParams = @{
                Name                = "CertIdentifier"
                CertificateFilePath = "C:\cert.cer"
                Ensure              = "Present"
            }

            Mock -CommandName Get-Item -MockWith {
                return @{
                    Subject    = "CN=CertName"
                    Thumbprint = $testParams.CertificateThumbprint
                }
            }

            Mock -CommandName Test-Path -MockWith {
                return $false
            }

            Mock -CommandName Get-SPTrustedRootAuthority -MockWith {
                return @{
                    Name        = $testParams.Name
                    Certificate = @{
                        Thumbprint = $testParams.CertificateThumbprint
                    }
                }
            }

            It "Should return Present from the Get method" {
                { Get-TargetResource @testParams } | Should Throw "Specified CertificateFilePath does not exist"
            }

            It "Should return true when the Test method is called" {
                { Test-TargetResource @testParams } | Should Throw "Specified CertificateFilePath does not exist"
            }

            It "Should Update the SP Trusted Root Authority in the set method" {
                { Set-TargetResource @testParams } | Should Throw "Specified CertificateFilePath does not exist"
            }
        }

        ## CertFile - RA does not exist

        Context -Name "When TrustedRootAuthority should exist and does exist in the farm (Thumbprint)." -Fixture {
            $testParams = @{
                Name                  = "CertIdentifier"
                CertificateThumbprint = "770515261D1AB169057E246E0EE6431D557C3AFB"
                Ensure                = "Present"
            }

            Mock -CommandName Get-Item -MockWith {
                return @{
                    Subject    = "CN=CertName"
                    Thumbprint = $testParams.CertificateThumbprint
                }
            }

            Mock -CommandName Get-SPTrustedRootAuthority -MockWith {
                return @{
                    Name        = $testParams.Name
                    Certificate = @{
                        Thumbprint = $testParams.CertificateThumbprint
                    }
                }
            }

            It "Should return Present from the Get method" {
                (Get-TargetResource @testParams).Ensure | Should Be "Present"
            }

            It "Should return true when the Test method is called" {
                Test-TargetResource @testParams | Should Be $true
            }

            It "Should Update the SP Trusted Root Authority in the set method" {
                Set-TargetResource @testParams
                Assert-MockCalled Get-SPTrustedRootAuthority -Times 1
                Assert-MockCalled Set-SPTrustedRootAuthority -Times 1
            }
        }

        Context -Name "When TrustedRootAuthority should exist and does exist in the farm (FilePath)." -Fixture {
            $testParams = @{
                Name                  = "CertIdentifier"
                CertificateThumbprint = "770515261D1AB169057E246E0EE6431D557C3AFB"
                Ensure                = "Present"
            }

            Mock -CommandName Test-Path -MockWith {
                return $true
            }

            Mock -CommandName Get-Item -MockWith {
                return  @{
                    Subject    = "CN=CertName"
                    Thumbprint = $testParams.CertificateThumbprint
                }
            }

            Mock -CommandName New-Object -MockWith {
                $retVal = [pscustomobject]@{
                    Subject       = "CN=CertIdentifer"
                    Thumbprint    = "770515261D1AB169057E246E0EE6431D557C3AFC"
                    HasPrivateKey = $false
                }
                Add-Member -InputObject $retVal -MemberType ScriptMethod Import { }

                return $retVal
            } -ParameterFilter { $TypeName -eq "System.Security.Cryptography.X509Certificates.X509Certificate2" }

            Mock -CommandName Get-SPTrustedRootAuthority -MockWith {
                return @{
                    Name        = $testParams.Name
                    Certificate = @{
                        Thumbprint = $testParams.CertificateThumbprint
                    }
                }
            }

            It "Should return Present from the Get method" {
                (Get-TargetResource @testParams).Ensure | Should Be "Present"
            }

            It "Should return true when the Test method is called" {
                Test-TargetResource @testParams | Should Be $true
            }

            It "Should Update the SP Trusted Root Authority in the set method" {
                Set-TargetResource @testParams
                Assert-MockCalled Get-SPTrustedRootAuthority -Times 1
                Assert-MockCalled Set-SPTrustedRootAuthority -Times 1
            }
        }

        Context -Name "When TrustedRootAuthority should exist and does exist in the farm, but has incorrect certificate (Thumbprint)." -Fixture {
            $testParams = @{
                Name                  = "CertIdentifier"
                CertificateThumbprint = "770515261D1AB169057E246E0EE6431D557C3AFB"
                Ensure                = "Present"
            }

            Mock -CommandName Get-SPTrustedRootAuthority -MockWith {
                return @{
                    Name        = $testParams.Name
                    Certificate = @{
                        Thumbprint = "770515261D1AB169057E246E0EE6431D557C3AFC"
                    }
                }
            }

            Mock -CommandName Get-Item -MockWith {
                return  @{
                    Subject    = "CN=CertName"
                    Thumbprint = $testParams.CertificateThumbprint
                }
            }

            It "Should return Present from the Get method" {
                (Get-TargetResource @testParams).Ensure | Should Be "Present"
            }

            It "Should return false when the Test method is called" {
                Test-TargetResource @testParams | Should Be $false
            }

            It "Should update the certificate in the set method" {
                Set-TargetResource @testParams
                Assert-MockCalled Get-SPTrustedRootAuthority -Times 1
                Assert-MockCalled Set-SPTrustedRootAuthority -Times 1
            }
        }

        Context -Name "When TrustedRootAuthority should exist and does exist in the farm, but has incorrect certificate (FilePath)." -Fixture {
            $testParams = @{
                Name                = "CertIdentifier"
                CertificateFilePath = "C:\cert.cer"
                Ensure              = "Present"
            }

            Mock -CommandName Get-SPTrustedRootAuthority -MockWith {
                return @{
                    Name        = $testParams.Name
                    Certificate = @{
                        Thumbprint = "770515261D1AB169057E246E0EE6431D557C3AFC"
                    }
                }
            }

            Mock -CommandName Test-Path -MockWith {
                return $true
            }

            Mock -CommandName Get-Item -MockWith {
                return  @{
                    Subject    = "CN=CertName"
                    Thumbprint = $testParams.CertificateThumbprint
                }
            }

            Mock -CommandName New-Object -MockWith {
                $retVal = [pscustomobject]@{
                    Subject       = "CN=CertIdentifer"
                    Thumbprint    = "770515261D1AB169057E246E0EE6431D557C3AFB"
                    HasPrivateKey = $false
                }
                Add-Member -InputObject $retVal -MemberType ScriptMethod Import { }

                return $retVal
            } -ParameterFilter { $TypeName -eq "System.Security.Cryptography.X509Certificates.X509Certificate2" }

            It "Should return Present from the Get method" {
                (Get-TargetResource @testParams).Ensure | Should Be "Present"
            }

            It "Should return false when the Test method is called" {
                Test-TargetResource @testParams | Should Be $false
            }

            It "Should update the certificate in the set method" {
                Set-TargetResource @testParams
                Assert-MockCalled Get-SPTrustedRootAuthority -Times 1
                Assert-MockCalled Set-SPTrustedRootAuthority -Times 1
            }
        }

        Context -Name "When TrustedRootAuthority should exist and does exist in the farm, but has incorrect certificate, but specified certificate doesn't exist;" -Fixture {
            $testParams = @{
                Name                  = "CertIdentifier"
                CertificateThumbprint = "770515261D1AB169057E246E0EE6431D557C3AFB"
                Ensure                = "Present"
            }

            Mock -CommandName Get-SPTrustedRootAuthority -MockWith {
                return @{
                    Name        = $testParams.Name
                    Certificate = @{
                        Thumbprint = "770515261D1AB169057E246E0EE6431D557C3AFC"
                    }
                }
            }

            Mock -CommandName Get-Item -MockWith {
                return  $null
            }

            It "Should return Present from the Get method" {
                (Get-TargetResource @testParams).Ensure | Should Be "Present"
            }

            It "Should return false when the Test method is called" {
                Test-TargetResource @testParams | Should Be $false
            }

            It "Should thorw Certificate not found error in the set method" {
                { Set-TargetResource @testParams } | Should Throw "Certificate not found in the local Certificate Store"
            }
        }

        Context -Name "When TrustedRootAuthority should exist and doesn't exist in the farm, but has an invalid certificate." -Fixture {
            $testParams = @{
                Name                  = "CertIdentifier"
                CertificateThumbprint = "770515261D1AB169057E246E0EE6431D557C3AFB"
                Ensure                = "Present"
            }

            Mock -CommandName Get-SPTrustedRootAuthority -MockWith {
                return $null
            }

            Mock -CommandName Get-Item -MockWith {
                return $null
            }

            It "Should return Absent from the Get method" {
                (Get-TargetResource @testParams).Ensure | Should Be "Absent"
            }

            It "Should return false when the Test method is called" {
                Test-TargetResource @testParams | Should Be $false
            }

            It "Should throw a Certificate not found error" {
                { Set-TargetResource @testParams } | Should Throw "Certificate not found in the local Certificate Store"
            }
        }

        Context -Name "When TrustedRootAuthority should exist and doesn't exist in the farm (Thumbprint)." -Fixture {
            $testParams = @{
                Name                  = "CertIdentifier"
                CertificateThumbprint = "770515261D1AB169057E246E0EE6431D557C3AFB"
                Ensure                = "Present"
            }

            Mock -CommandName Get-Item -MockWith {
                return @{
                    Subject    = "CN=CertIdentifier"
                    Thumbprint = $testParams.CertificateThumbprint
                }
            }

            Mock -CommandName Get-SPTrustedRootAuthority -MockWith {
                return $null
            }

            It "Should return Absent from the Get method" {
                (Get-TargetResource @testParams).Ensure | Should Be "Absent"
            }

            It "Should return true when the Test method is called" {
                Test-TargetResource @testParams | Should Be $false
            }

            It "Should create a new service application in the set method" {
                Set-TargetResource @testParams
                Assert-MockCalled Get-Item -Times 1
                Assert-MockCalled New-SPTrustedRootAuthority -Times 1
            }
        }

        Context -Name "When TrustedRootAuthority should exist and doesn't exist in the farm (FilePath)." -Fixture {
            $testParams = @{
                Name                = "CertIdentifier"
                CertificateFilePath = "c:\cert.cer"
                Ensure              = "Present"
            }

            Mock -CommandName Test-Path -MockWith {
                return $true
            }

            Mock -CommandName Get-Item -MockWith {
                return  @{
                    Subject    = "CN=CertName"
                    Thumbprint = $testParams.CertificateThumbprint
                }
            }

            Mock -CommandName New-Object -MockWith {
                $retVal = [pscustomobject]@{
                    Subject       = "CN=CertIdentifer"
                    Thumbprint    = "770515261D1AB169057E246E0EE6431D557C3AFB"
                    HasPrivateKey = $false
                }
                Add-Member -InputObject $retVal -MemberType ScriptMethod Import { }

                return $retVal
            } -ParameterFilter { $TypeName -eq "System.Security.Cryptography.X509Certificates.X509Certificate2" }

            Mock -CommandName Get-SPTrustedRootAuthority -MockWith {
                return $null
            }

            It "Should return Absent from the Get method" {
                (Get-TargetResource @testParams).Ensure | Should Be "Absent"
            }

            It "Should return true when the Test method is called" {
                Test-TargetResource @testParams | Should Be $false
            }

            It "Should create a new service application in the set method" {
                Set-TargetResource @testParams
                Assert-MockCalled New-SPTrustedRootAuthority -Times 1
            }
        }

        Context -Name "When TrustedRootAuthority should exist and doesn't exist in the farm, but specified cert contains a private key" -Fixture {
            $testParams = @{
                Name                  = "CertIdentifier"
                CertificateThumbprint = "770515261D1AB169057E246E0EE6431D557C3AFB"
                Ensure                = "Present"
            }

            Mock -CommandName Get-Item `
                -MockWith {
                $retVal = [pscustomobject]@{
                    Subject       = "CN=CertIdentifier"
                    Thumbprint    = $testParams.CertificateThumbprint
                    HasPrivateKey = $true
                }

                Add-Member -InputObject $retVal -MemberType ScriptMethod Export {
                    $bytes = [System.Byte[]]::CreateInstance([System.Byte], 512)
                    return $bytes
                }

                return $retVal
            }

            Mock -CommandName New-Object `
                -ParameterFilter { $TypeName -eq "System.Security.Cryptography.X509Certificates.X509Certificate2" } `
                -MockWith {
                $retVal = [pscustomobject]@{ }
                Add-Member -InputObject $retVal -MemberType ScriptMethod Import {
                    param([System.Byte[]]$bytes)
                    return @{
                        Subject       = "CN=CertIdentifer"
                        Thumbprint    = $testParams.CertificateThumbprint
                        HasPrivateKey = $false
                    }
                }

                return $retVal
            }

            Mock -CommandName Get-SPTrustedRootAuthority -MockWith {
                return $null
            }

            It "Should return Absent from the Get method" {
                (Get-TargetResource @testParams).Ensure | Should Be "Absent"
            }

            It "Should return true when the Test method is called" {
                Test-TargetResource @testParams | Should Be $false
            }

            It "Should create a new Trusted Root Authority in the set method" {
                Set-TargetResource @testParams
                Assert-MockCalled Get-Item -Times 1
                Assert-MockCalled New-SPTrustedRootAuthority -Times 1
                Assert-MockCalled New-Object -Times 1
            }
        }

        Context -Name "When TrustedRootAuthority should exist and does exist but is incorrect certificate and specified cert contains a private key" -Fixture {
            $testParams = @{
                Name                  = "CertIdentifier"
                CertificateThumbprint = "770515261D1AB169057E246E0EE6431D557C3AFB"
                Ensure                = "Present"
            }

            Mock -CommandName Get-Item -MockWith {
                $retVal = [pscustomobject]@{
                    Subject       = "CN=CertIdentifier"
                    Thumbprint    = $testParams.CertificateThumbprint
                    HasPrivateKey = $true
                }

                Add-Member -InputObject $retVal -MemberType ScriptMethod Export {
                    $bytes = [System.Byte[]]::CreateInstance([System.Byte], 512)
                    return $bytes
                }

                return $retVal
            }

            Mock -CommandName New-Object -MockWith {
                $retVal = [pscustomobject]@{ }
                Add-Member -InputObject $retVal -MemberType ScriptMethod Import {
                    param([System.Byte[]]$bytes)
                    return @{
                        Subject       = "CN=CertIdentifer"
                        Thumbprint    = $testParams.CertificateThumbprint
                        HasPrivateKey = $false
                    }
                }

                return $retVal
            } -ParameterFilter { $TypeName -eq "System.Security.Cryptography.X509Certificates.X509Certificate2" }

            Mock -CommandName Get-SPTrustedRootAuthority -MockWith {
                return @{
                    Name        = $testParams.Name
                    Certificate = @{
                        Thumbprint = "770515261D1AB169057E246E0EE6431D557C3AFC"
                    }
                }
            }

            It "Should return Absent from the Get method" {
                (Get-TargetResource @testParams).Ensure | Should Be "Present"
            }

            It "Should return true when the Test method is called" {
                Test-TargetResource @testParams | Should Be $false
            }

            It "Should create a new Trusted Root Authority in the set method" {
                Set-TargetResource @testParams
                Assert-MockCalled Get-Item -Times 1
                Assert-MockCalled Set-SPTrustedRootAuthority -Times 1
                Assert-MockCalled New-Object -Times 1
            }
        }

        Context -Name "When TrustedRootAuthority shouldn't exist and does exist in the farm." -Fixture {
            $testParams = @{
                Name                  = "CertIdentifier"
                CertificateThumbprint = "770515261D1AB169057E246E0EE6431D557C3AFB"
                Ensure                = "Absent"
            }

            Mock -CommandName Get-Item -MockWith {
                return @{
                    Subject    = "CN=CertIdentifier"
                    Thumbprint = $testParams.CertificateThumbprint
                }
            }

            Mock -CommandName Get-SPTrustedRootAuthority -MockWith {
                return @{
                    Name        = $testParams.Name
                    Certificate = @{
                        Thumbprint = $testParams.CertificateThumbprint
                    }
                }
            }

            It "Should return Present from the Get method" {
                (Get-TargetResource @testParams).Ensure | Should Be "Present"
            }

            It "Should return false when the Test method is called" {
                Test-TargetResource @testParams | Should Be $false
            }

            It "Should remove the Trusted Root Authority" {
                Set-TargetResource @testParams
                Assert-MockCalled Remove-SPTrustedRootAuthority -Times 1
            }
        }

        Context -Name "When TrustedRootAuthority shouldn't exist and doesn't exist in the farm." -Fixture {
            $testParams = @{
                Name                  = "CertIdentifier"
                CertificateThumbprint = "770515261D1AB169057E246E0EE6431D557C3AFB"
                Ensure                = "Absent"
            }

            Mock -CommandName Get-Item -MockWith {
                return  @{
                    Subject    = "CN=CertIdentifier"
                    Thumbprint = $testParams.CertificateThumbprint
                }
            }

            Mock -CommandName Get-SPTrustedRootAuthority -MockWith {
                return $null
            }

            It "Should return Absent from the Get method" {
                (Get-TargetResource @testParams).Ensure | Should Be "Absent"
            }

            It "Should return false when the Test method is called" {
                Test-TargetResource @testParams | Should Be $true
            }

            It "Should remove the Trusted Root Authority" {
                Set-TargetResource @testParams
                Assert-MockCalled Remove-SPTrustedRootAuthority -Times 1
            }
        }
    }
}
