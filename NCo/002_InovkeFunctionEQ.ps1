
#-Begin-----------------------------------------------------------------

  <#
    .Synopsis
      Ruft einen Easy Query Funktionsbaustein auf.
    .Description
      Ruft den Easy Query Funktionsbaustein /BIC/NF_13 auf, übergibt die
      Importparameter und zeigt das Ergebnis von E_T_COLUMN_DESCRIPTION
      an.
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
        
        $cfgParams.Add($cfgParams::Name, "Test")
        $cfgParams.Add($cfgParams::AppServerHost, "ABAP")
        $cfgParams.Add($cfgParams::SystemNumber, "00")
        $cfgParams.Add($cfgParams::Client, "001")
        $cfgParams.Add($cfgParams::User, "BCUSER")
        $cfgParams.Add($cfgParams::Password, "minisap")
        $cfgParams.Add($cfgParams::UseSAPGui, "2")
        $cfgParams.Add($cfgParams::AbapDebug, "1")

      Return [SAP.Middleware.Connector.RfcDestinationManager]::GetDestination($cfgParams)

    }
  
  #-Sub Invoke-SAPFunctionModule----------------------------------------
    Function Invoke-SAPFunctionModule {

      $destination = Get-Destination

      #-Metadaten-------------------------------------------------------
        $rfcFunction = `
          $destination.Repository.CreateFunction("/BIC/NF_13")

      #-Importparameter setzen------------------------------------------
        [SAP.Middleware.Connector.IRfcStructure]$imCM91VJAHR = `
          $rfcFunction.GetStructure("I_S_VAR_01CM91VJAHR")
        $imCM91VJAHR.SetValue("SIGN", "I")
        $imCM91VJAHR.SetValue("OPTION", "EQ")
        $imCM91VJAHR.SetValue("LOW", "2013")

        [SAP.Middleware.Connector.IRfcStructure]$imCM91VBEREICHQ = `
          $rfcFunction.GetStructure("I_S_VAR_02CM91VBEREICHQ")
        $imCM91VBEREICHQ.SetValue("SIGN", "I")
        $imCM91VBEREICHQ.SetValue("OPTION", "EQ")
        $imCM91VBEREICHQ.SetValue("LOW", "MGV")

      #-Funktionsbaustein aufrufen--------------------------------------
        Try {
          $rfcFunction.Invoke($destination)
        }
        Catch [SAP.Middleware.Connector.RfcCommunicationException], `
          [SAP.Middleware.Connector.RfcLogonException], `
          [SAP.Middleware.Connector.RfcAbapRuntimeException] {
            Write-Host "Exception: " $_.Exception.Message
            Break
          }
        Catch [SAP.Middleware.Connector.RfcAbapClassException] {
          Write-Host "Exception: " $_.Excepton.Message
          Break
        }
        Catch [SAP.Middleware.Connector.RfcAbapException] {
          Write-Host "Exception: " $_.Exception.Message
          Break
        }
        Catch [SAP.Middleware.Connector.RfcAbapMessageException] {
          Write-Host "Exception: " $_.Exception.Message
          Break
        }

      #-Exportparameter anzeigen----------------------------------------
        [SAP.Middleware.Connector.IRfcTable]$exCOLUMNDESCR = `
          $rfcFunction.GetTable("E_T_COLUMN_DESCRIPTION")
        ForEach ($line in $exCOLUMNDESCR) {
          Write-Host $line.GetValue("COLUMN_NAME")
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
