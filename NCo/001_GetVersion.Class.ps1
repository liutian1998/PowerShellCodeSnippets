
#-Begin-----------------------------------------------------------------

  <#
    .Synopsis
      Zeigt die Version des NCo an.
    .Description
      Zeigt die Versionsnummer, das Patchlevel und das SAP-Release
      des genutzten dotNET Connectors (NCo) für SAP an.
  #>

  #-Class NCo-----------------------------------------------------------
    Class NCo {

      #-Constructor-----------------------------------------------------
      #-
      #- Das Laden der NCo-Bibliotheken erfolgt in Abhaengigkeit zur
      #- eingesetzten Shell. Bei 64-bit werden die Bibliotheken aus dem
      #- Pfad x64 geladen, bei 32-bit aus dem Pfad x86.
      #-
      #-----------------------------------------------------------------
      NCo() {

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

      #-Method GetVersion-----------------------------------------------
      [Void]GetVersion() {

        #-Version des NCo anzeigen--------------------------------------
          $ConnInfo = New-Object SAP.Middleware.Connector.SAPConnectorInfo
          $Version = $ConnInfo::get_Version()
          $PatchLevel = $ConnInfo::get_KernelPatchLevel()
          $SAPRelease = $ConnInfo::get_SAPRelease()

        Write-Host "`r`nNCo verion:" $Version
        Write-Host "Patch Level:" $PatchLevel
        Write-Host "SAP Release:" $SAPRelease

      }

    }

  #-Sub Main------------------------------------------------------------
    Function Main () {
      If ($PSVersionTable.PSVersion.Major -ge 5) {
        $NCo = [NCo]::new()
        $NCo.GetVersion()
      }
    }

  #-Main----------------------------------------------------------------
    Main

#-End-------------------------------------------------------------------
