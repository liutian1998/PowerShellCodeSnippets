
#-Begin-----------------------------------------------------------------

  <#
    .Synopsis
      Calls RFC_SYSTEM_INFO of an SAP system.
    .Description
      Calls the function module RFC_SYSTEM_INFO and writes the result
      on the screen.
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
          $File = $Path + "sapnco_utils.dll"; Add-Type -Path $File

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

      #-Method GetSystemInfo--------------------------------------------
        GetSystemInfo() {

          $destination = $This.GetDestination()

          #-Metadata----------------------------------------------------
            $rfcFunction = `
              $destination.Repository.CreateFunction("RFC_SYSTEM_INFO")

          #-Call function module----------------------------------------
            Try {
              $rfcFunction.Invoke($destination)

              $Export = $rfcFunction.GetStructure("RFCSI_EXPORT")

              #-Get information-----------------------------------------
                Write-Host $Export.GetValue("RFCHOST")
                Write-Host $Export.GetValue("RFCSYSID")
                Write-Host $Export.GetValue("RFCDBHOST")
                Write-Host $Export.GetValue("RFCDBSYS")

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
        $NCo.GetSystemInfo()
      }
    }

  #-Main----------------------------------------------------------------
    Main

#-End-------------------------------------------------------------------
