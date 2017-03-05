
#-Begin-----------------------------------------------------------------

  <#
    .Synopsis
      Ruft den Funktionsbaustein BAPI_USER_DELETE auf.
    .Description
      Ruft den Funktionsbaustein BAPI_USER_DELETE auf, übergibt die
      Importparameter und loescht den User MYUSER an.
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

      [Reflection.Assembly]::LoadFile($Path + "sapnco.dll") > $Null
      [Reflection.Assembly]::LoadFile($Path + "sapnco_utils.dll") > $Null

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

      #-Metadaten-------------------------------------------------------
        Try {
          [SAP.Middleware.Connector.IRfcFunction]$bapiDeleteUser = 
            $destination.Repository.CreateFunction("BAPI_USER_DELETE")
          [SAP.Middleware.Connector.IRfcFunction]$bapiTransactionCommit = 
            $destination.Repository.CreateFunction("BAPI_TRANSACTION_COMMIT")
        }
        Catch [SAP.Middleware.Connector.RfcBaseException] {
          Write-Host "Metadaten-Fehler"
          Break
        }

      #-Importparameter setzen------------------------------------------
        $bapiDeleteUser.SetValue("USERNAME", "MYUSER")

      #-Aufrufkontext oeffnen-------------------------------------------
        [SAP.Middleware.Connector.RfcSessionManager]::BeginContext($destination) > $Null

      #-Funktionsbausteine ausfuehren-----------------------------------
        Try {
          #-User loeschen-----------------------------------------------
            $bapiDeleteUser.Invoke($destination)
          [SAP.Middleware.Connector.IRfcTable]$return = 
            $bapiDeleteUser.GetTable("RETURN")
          ForEach ($line in $return) {
            Write-Host $line.GetValue("TYPE") "-" $line.GetValue("MESSAGE")
          }
          #-Commit durchfuehren-----------------------------------------
            $bapiTransactionCommit.Invoke($destination)
        }
        Finally {
          #-Aufrufkontext schliessen------------------------------------
            [SAP.Middleware.Connector.RfcSessionManager]::EndContext($destination) > $Null
        }

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
