
#-Begin-----------------------------------------------------------------

  <#
    .Synopsis
      Ruft den Funktionsbaustein STFC_CONNECTION auf.
    .Description
      Ruft den Funktionsbaustein STFC_CONNECTION auf, übergibt einen
      Importparameter und zeigt das Ergebnis der Exportparameter an.
      Das Passwort wird mittels eins gesicherten Eingabefeldes
      uebergeben, der erfolgreiche Verbindungsaufbau wird geprueft und
      beim Aufruf des Funktionsbausteines werden Ausnahmen abgefangen.
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

  #-Sub Invoke-SAPFunctionModule----------------------------------------
    Function Invoke-SAPFunctionModule {

      $destination = Get-Destination
      Try {
        $rc = $destination.Repository.ToString()
      }
      Catch {
        Write-Host "Verbindungsfehler"
        Break
      }

      #-Metadaten-------------------------------------------------------
        $rfcFunction = `
          $destination.Repository.CreateFunction("STFC_CONNECTION")

      #-Importparameter setzen------------------------------------------
        $rfcFunction.SetValue("REQUTEXT", "Hello World from PowerShell")

      #-Funktionsbaustein aufrufen und Ausnahmen abfangen---------------
        Try {
          $rfcFunction.Invoke($destination)
        }
        Catch [SAP.Middleware.Connector.RfcCommunicationException], `
          [SAP.Middleware.Connector.RfcLogonException], `
          [SAP.Middleware.Connector.RfcAbapRuntimeException] {
          Write-Host "Exception occured:`r`n" $_.Exception.Message
          Break
        }
        Catch [SAP.Middleware.Connector.RfcAbapClassException] {
          Write-Host "Exception occured:`r`n" $_.Excepton.Message
          Break
        }
        Catch [SAP.Middleware.Connector.RfcAbapException] {
          Write-Host "Exception occured:`r`n" $_.Exception.Message
          Break
        }
        Catch [SAP.Middleware.Connector.RfcAbapMessageException] {
          Write-Host "Exception occured:`r`n" $_.Exception.Message
          Break
        }

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
