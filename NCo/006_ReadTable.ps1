
#-Begin-----------------------------------------------------------------

  <#
    .Synopsis
      Calls RFC_READ_TABLE of an SAP system.
    .Description
      Calls the function module RFC_READ_TABLE and writes the result
      on the screen.
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
  
  #-Sub Read-Table------------------------------------------------------
    Function Read-Table ([String]$TableName) {

      $destination = Get-Destination

      #-Metadata--------------------------------------------------------
        $rfcFunction = `
          $destination.Repository.CreateFunction("RFC_READ_TABLE")

      #-Call function module--------------------------------------------
        Try {
          $rfcFunction.SetValue("QUERY_TABLE", $TableName)
          $rfcFunction.SetValue("DELIMITER", ";")

          $rfcFunction.Invoke($destination)

          [SAP.Middleware.Connector.IRfcTable]$Table = `
            $rfcFunction.GetTable("DATA")

          ForEach ($Line In $Table) {
            Write-Host $Line.GetValue("WA")
          }

        }
        Catch {
          Write-Host "Exception occured:`r`n" $_.Exception.Message
        }


    }

  #-Sub Main------------------------------------------------------------
    Function Main () {
      If ($PSVersionTable.PSVersion.Major -ge 5) {
        Load-NCo
        Read-Table "USR01"
      }
    }

  #-Main----------------------------------------------------------------
    Main

#-End-------------------------------------------------------------------
