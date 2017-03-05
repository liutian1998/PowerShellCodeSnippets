
#-Begin-----------------------------------------------------------------

  <#
    .Synopsis
      Ruft den Funktionsbaustein BAPI_USER_CREATE1 auf.
    .Description
      Ruft den Funktionsbaustein BAPI_USER_CREATE1 auf, übergibt die
      Importparameter und legt den User MYUSER an.
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
          [SAP.Middleware.Connector.IRfcFunction]$bapiCreateUser = 
            $destination.Repository.CreateFunction("BAPI_USER_CREATE1")
          [SAP.Middleware.Connector.IRfcFunction]$bapiTransactionCommit = 
            $destination.Repository.CreateFunction("BAPI_TRANSACTION_COMMIT")
        }
        Catch [SAP.Middleware.Connector.RfcBaseException] {
          Write-Host "Metadaten-Fehler"
          Break
        }

      #-Importparameter setzen------------------------------------------
        $bapiCreateUser.SetValue("USERNAME", "MYUSER")
        [SAP.Middleware.Connector.IRfcStructure]$imPassword = 
          $bapiCreateUser.GetStructure("PASSWORD")
        $imPassword.SetValue("BAPIPWD", "initial")
        [SAP.Middleware.Connector.IRfcStructure]$imAddress = 
          $bapiCreateUser.GetStructure("ADDRESS")
        $imAddress.SetValue("FIRSTNAME", "My")
        $imAddress.SetValue("LASTNAME", "User")
        $imAddress.SetValue("FULLNAME", "MyUser")

      #-Aufrufkontext oeffnen-------------------------------------------
        [SAP.Middleware.Connector.RfcSessionManager]::BeginContext($destination) > $Null

      #-Funktionsbausteine ausfuehren-----------------------------------
        Try {
          #-User anlegen------------------------------------------------
            $bapiCreateUser.Invoke($destination)
          [SAP.Middleware.Connector.IRfcTable]$return = 
            $bapiCreateUser.GetTable("RETURN")
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
