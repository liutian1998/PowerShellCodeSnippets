
#-Begin-----------------------------------------------------------------

  <#
    .Synopsis
      Ruft den Funktionsbaustein STFC_CONNECTION auf.
    .Description
      Ruft den Funktionsbaustein STFC_CONNECTION auf, übergibt einen
      Importparameter und zeigt das Ergebnis der Exportparameter an.
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

      #-Verbindungsparamter---------------------------------------------
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

      Return [SAP.Middleware.Connector.RfcDestinationManager]::GetDestination($cfgParams)

    }
  
  #-Sub Invoke-SAPFunctionModule----------------------------------------
    Function Invoke-SAPFunctionModule {

      $destination = Get-Destination

      #-Metadaten-------------------------------------------------------
        $rfcFunction = `
          $destination.Repository.CreateFunction("STFC_CONNECTION")

      #-Importparameter setzen------------------------------------------
        $rfcFunction.SetValue("REQUTEXT", "Hello World from PowerShell")

      #-Funktionsbaustein aufrufen--------------------------------------
        $rfcFunction.Invoke($destination)

      #-Exportparameter anzeigen----------------------------------------
        Write-Host $rfcFunction.GetValue("ECHOTEXT")
        Write-Host $rfcFunction.GetValue("RESPTEXT")

    }

  #-Sub Main------------------------------------------------------------
    Function Main () {
      If ($PSVersionTable.PSVersion.Major -ge 5) {
        Load-NCo
        Invoke-SAPFunctionModule
      }
    }

  #-Main----------------------------------------------------------------
    Main

#-End-------------------------------------------------------------------
