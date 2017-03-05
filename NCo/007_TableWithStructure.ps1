
#-Begin-----------------------------------------------------------------
#-
#- Hint: The TimeOut of an RFC call can be set via the transaction
#- code RZ11 with the profile parameter rdisp/max_wprun_time.
#-
#-----------------------------------------------------------------------

  <#
    .Synopsis
      Shows an example how to use a table in a structure.
    .Description
      Shows an example how to use a table in a structure of an
      import parameter of an RFC-enabled function module.

      Local interface of the function module Z_INFO_EXT:
      Importing
        Value(IS_IN) Type /GKV/CM91_STR_BS_HZV_EFN_IN
      Exporting
        Value(ET_OUT) Type /GKV/CM91_TAB_BS_HZV_EFN_OUT
        Value(RETURN) Type BAPIRET2_T

      The structure IS_IN contains a table TAB_KEY from the type
      /BOBF/T_FRW_KEY2.
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

      #-Connectionparameter---------------------------------------------
        $cfgParams = New-Object SAP.Middleware.Connector.RfcConfigParameters
        $cfgParams.Add("NAME", "TEST")
        $cfgParams.Add("ASHOST", "ABAP")
        $cfgParams.Add("SYSNR", "00")
        $cfgParams.Add("CLIENT", "001")
        $cfgParams.Add("USER", "BCUSER")
        #$cfgParams.Add("USE_SAPGUI", "2")
        #$cfgParams.Add("ABAP_DEBUG", "1")

        $SecPasswd = Read-Host -Prompt "Passwort" -AsSecureString
        $ptrPasswd = [Runtime.InteropServices.Marshal]::SecureStringToBStr($SecPasswd)
        $Passwd = [Runtime.InteropServices.Marshal]::PtrToStringBStr($ptrPasswd)
        $cfgParams.Add("PASSWD", $Passwd)

      Return [SAP.Middleware.Connector.RfcDestinationManager]::GetDestination($cfgParams)

    }
  
  #-Sub Invoke-SAPFunctionModule----------------------------------------
    Function Invoke-SAPFunctionModule {

      $destination = Get-Destination

      #-Metadata--------------------------------------------------------
        [SAP.Middleware.Connector.IRfcFunction]$rfcFunction = `
          $destination.Repository.CreateFunction("Z_INFO_EXT")

      #-Get import parameter--------------------------------------------
        [SAP.Middleware.Connector.IRfcStructure]$is_in = `
          $rfcFunction.GetStructure("IS_IN")

      #-Get type of table-----------------------------------------------
        [SAP.Middleware.Connector.RfcTableMetadata]$tabMetadata = `
          $destination.Repository.GetTableMetadata("/BOBF/T_FRW_KEY2")

      #-Create a table of type /BOBF/T_FRW_KEY2-------------------------
        [SAP.Middleware.Connector.IRfcTable]$tab_key = `
          $tabMetadata.CreateTable()

      #-Read keys from CSV file-----------------------------------------
        $Keys = Import-Csv "lt_keys.csv" -Delimiter ";"

      #-Copy keys from CSV file to table--------------------------------
        ForEach($Key in $Keys) {
          $tab_key.Append()
          $tab_key.SetValue(0, $Key.DB_KEY)
        }

      #-Set import parameter--------------------------------------------
        $is_in.SetValue("TAB_KEY" , $tab_key)

      #-Call function module--------------------------------------------
        Try {

          $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

            $rfcFunction.Invoke($destination)

          $Stopwatch.Stop()

          Write-Host $Stopwatch.Elapsed.TotalMilliseconds `
            "Milliseconds with" $tab_key.Count "keys"

          #-Show export parameter---------------------------------------
            [SAP.Middleware.Connector.IRfcTable]$Table = `
              $rfcFunction.GetTable("ET_OUT")

        }
        Catch {
          Write-Host "Exception occured:`r`n" $_.Exception.Message
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
