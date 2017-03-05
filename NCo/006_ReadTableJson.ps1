
#-Begin-----------------------------------------------------------------

  <#
    .Synopsis
      Calls RFC_READ_TABLE of an SAP system.
    .Description
      Calls the function module RFC_READ_TABLE and writes the result
      in JSON on the screen.
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
  
  #-Function Read-Table-------------------------------------------------
    Function Read-Table-JSON ([String]$TableName, [Int]$RowCount) {

      $destination = Get-Destination

      #-Metadata--------------------------------------------------------
        $rfcFunction = `
          $destination.Repository.CreateFunction("RFC_READ_TABLE")

      #-Call function module--------------------------------------------
        Try {
          $rfcFunction.SetValue("QUERY_TABLE", $TableName)
          $rfcFunction.SetValue("DELIMITER", "~")
          $rfcFunction.SetValue("ROWCOUNT", $RowCount)

          $rfcFunction.Invoke($destination)

          $Header = New-Object System.Collections.ArrayList

          #-Get field names for the CSV header--------------------------
            [SAP.Middleware.Connector.IRfcTable]$Fields = `
              $rfcFunction.GetTable("FIELDS")
            ForEach ($Field In $Fields) {
              [Void]$Header.Add($Field.GetValue("FIELDNAME"))
            }

          #-Get table data line by line for CSV-------------------------
            [SAP.Middleware.Connector.IRfcTable]$Table = `
              $rfcFunction.GetTable("DATA")
            ForEach ($Line In $Table) {
              $CSV = $CSV + $Line.GetValue("WA") + "`r`n"
            }

          #-Convert CSV to JSON-----------------------------------------
            $JSON = $CSV | ConvertFrom-Csv -Delimiter "~" `
              -Header $Header | ConvertTo-Json

          Return $JSON

        }
        Catch {
          Write-Host "Exception occured:`r`n" $_.Exception.Message
        }


    }

  #-Sub Main------------------------------------------------------------
    Function Main () {
      If ($PSVersionTable.PSVersion.Major -ge 5) {
        Load-NCo
        Write-Host(Read-Table-JSON "TFDIR" 10)
      }
    }

  #-Main----------------------------------------------------------------
    Main

#-End-------------------------------------------------------------------
