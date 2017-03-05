
#-Begin-----------------------------------------------------------------

  <#
    .Synopsis
      Ruft den Funktionsbaustein STFC_CONNECTION auf.
    .Description
      Ruft den Funktionsbaustein STFC_CONNECTION auf, übergibt einen
      Importparameter und zeigt das Ergebnis der Exportparameter an.
  #>

  <#
    Hint:
    Every PowerShell script is completely parsed before the first
    statement in the script is executed. An unresolvable type name
    token inside a class definition is considered a parse error. To
    solve this problem, you have to load your types before the class
    definition is parsed.
  #>

  #-Function Get-Destination--------------------------------------------
    Function Get-Destination($cfgParams) {
      If (("SAP.Middleware.Connector.RfcDestinationManager" -as [type]) -ne $null) {
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

      #-Method InvokeSAPFunctionModule----------------------------------
        InvokeSAPFunctionModule() {

          $destination = $This.GetDestination()

          #-Metadaten---------------------------------------------------
            $rfcFunction = `
              $destination.Repository.CreateFunction("STFC_CONNECTION")

          #-Importparameter setzen--------------------------------------
            $rfcFunction.SetValue("REQUTEXT", "Hello World from PowerShell")

          #-Funktionsbaustein aufrufen----------------------------------
            $rfcFunction.Invoke($destination)

          #-Exportparameter anzeigen------------------------------------
            Write-Host $rfcFunction.GetValue("ECHOTEXT")
            Write-Host $rfcFunction.GetValue("RESPTEXT")

        }

    }

  #-Sub Main------------------------------------------------------------
    Function Main () {
      If ($PSVersionTable.PSVersion.Major -ge 5) {
        $NCo = [NCo]::new()
        $NCo.InvokeSAPFunctionModule()
      }
    }

  #-Main----------------------------------------------------------------
    Main

#-End-------------------------------------------------------------------
