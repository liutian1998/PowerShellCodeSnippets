
#-Begin-----------------------------------------------------------------

  <#
    .Synopsis
      Calls PING of an SAP system.
    .Description
      Calls the function module RFC_PING and the build-in function
      Ping of the .NET connector.
  #>

  #-Function Get-Destination--------------------------------------------
    Function Get-Destination($cfgParams) {
      If (("SAP.Middleware.Connector.RfcDestinationManager" -as [type]) -ne $null){
        Return [SAP.Middleware.Connector.RfcDestinationManager]::GetDestination($cfgParams)
      }
    }

  #-Class NCo-----------------------------------------------------------
    Class NCo {

      #-Constructor-----------------------------------------------------
        NCo() {

          [String]$ScriptDir = $PSScriptRoot

          If ([Environment]::Is64BitProcess) {
            [String]$Path = $ScriptDir + "\x64\"
          }
          Else {
            [String]$Path = $ScriptDir + "\x86\"
          }

          [String]$File = $Path + "sapnco.dll"; Add-Type -Path $File
          $File = $Path + "sapnco_utils.dll" ; Add-Type -Path $File

        }

      #-Method GetDestination-------------------------------------------
        Hidden [Object]GetDestination() {

          #-Verbindungsparamter-----------------------------------------
            $cfgParams = `
              New-Object SAP.Middleware.Connector.RfcConfigParameters
            $cfgParams.Add($cfgParams::Name, "TEST")
            $cfgParams.Add($cfgParams::AppServerHost, "ABAP")
            $cfgParams.Add($cfgParams::SystemNumber, "00")
            $cfgParams.Add($cfgParams::Client, "001")
            $cfgParams.Add($cfgParams::User, "BCUSER")

            $SecPasswd = Read-Host -Prompt "Passwort" -AsSecureString
            $ptrPasswd = `
              [Runtime.InteropServices.Marshal]::SecureStringToBStr($SecPasswd)
            $Passwd = `
              [Runtime.InteropServices.Marshal]::PtrToStringBStr($ptrPasswd)
            $cfgParams.Add($cfgParams::Password, $Passwd)

          Return Get-Destination($cfgParams)

        }

      #-Method ExecutePing----------------------------------------------
        ExecutePing() {

          $destination = $This.GetDestination()

          #-Metadata----------------------------------------------------
            $rfcFunction = `
              $destination.Repository.CreateFunction("RFC_PING")

          #-Variant 1: Call function module-----------------------------
            Try {
              $rfcFunction.Invoke($destination)
              Write-Host "Ping successful"
            }
            Catch {
              Write-Host "Exception occured:`r`n" $_.Exception.Message
            }

          #-Variant 2: Call build-in function---------------------------
            Try {
              $destination.Ping()
              Write-Host "Ping successful"
            }
            Catch {
              Write-Host "Exception occured:`r`n" $_.Exception.Message
            }

        }

    }

  #-Sub Main------------------------------------------------------------
    Function Main () {
      If ($PSVersionTable.PSVersion.Major -ge 5) {
        $NCo = [NCo]::new()
        $NCo.ExecutePing()
      }
    }

  #-Main----------------------------------------------------------------
    Main

#-End-------------------------------------------------------------------
