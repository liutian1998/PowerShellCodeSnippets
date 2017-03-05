
#-Begin-----------------------------------------------------------------

  <#
    .Synopsis
      Calls PING of an SAP system.
    .Description
      Calls the function module RFC_PING and the build-in function
      Ping of the .NET connector.
  #>

  #-Sub Load-NCo--------------------------------------------------------
    Function Load-NCo {

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

  #-Function Get-Destination--------------------------------------------
    Function Get-Destination {

      #-Connection parameters-------------------------------------------
        $cfgParams = New-Object SAP.Middleware.Connector.RfcConfigParameters
        $cfgParams.Add("NAME", "TEST")
        $cfgParams.Add("ASHOST", "ABAP")
        $cfgParams.Add("SYSNR", "00")
        $cfgParams.Add("CLIENT", "001")
        $cfgParams.Add("USER", "BCUSER")

        $SecPasswd = Read-Host -Prompt "Passwort" -AsSecureString
        $ptrPasswd = [Runtime.InteropServices.Marshal]::SecureStringToBStr($SecPasswd)
        $Passwd = [Runtime.InteropServices.Marshal]::PtrToStringBStr($ptrPasswd)
        $cfgParams.Add("PASSWD", $Passwd)

      Return [SAP.Middleware.Connector.RfcDestinationManager]::GetDestination($cfgParams)

    }
  
  #-Sub Execute-Ping----------------------------------------------------
    Function Execute-Ping () {

      $destination = Get-Destination

      #-Metadata--------------------------------------------------------
        $rfcFunction = `
          $destination.Repository.CreateFunction("RFC_PING")

      #-Variant 1: Call function module---------------------------------
        Try {
          $rfcFunction.Invoke($destination)
          Write-Host "Ping successful"
        }
        Catch {
          Write-Host "Exception occured:`r`n" $_.Exception.Message
        }

      #-Variant 2: Call build-in function-------------------------------
        Try {
          $destination.Ping()
          Write-Host "Ping successful"
        }
        Catch {
          Write-Host "Exception occured:`r`n" $_.Exception.Message
        }

    }

  #-Sub Main------------------------------------------------------------
    Function Main () {
      If ($PSVersionTable.PSVersion.Major -ge 5) {
        Load-NCo
        Execute-Ping
      }
    }

  #-Main----------------------------------------------------------------
    Main

#-End-------------------------------------------------------------------
