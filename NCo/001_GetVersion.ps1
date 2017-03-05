
#-Begin-----------------------------------------------------------------

  <#
    .Synopsis
      Zeigt die Version des NCo an.
    .Description
      Zeigt die Versionsnummer, das Patchlevel und das SAP-Release
      des genutzten dotNET Connectors (NCo) für SAP an.
  #>

  #-Sub Load-NCo--------------------------------------------------------
  #-
  #- Das Laden der NCo-Bibliotheken erfolgt in Abhaengigkeit zur
  #- eingesetzten Shell. Bei 64-bit werden die Bibliotheken aus dem 
  #- Pfad x64 geladen, bei 32-bit aus dem Pfad x86.
  #-
  #---------------------------------------------------------------------
    Function Load-NCo () {

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

  #-Sub Get-NCoVersion--------------------------------------------------
    Function Get-NCoVersion () {

      #-Version des NCo anzeigen----------------------------------------
        $Version = 
          [SAP.Middleware.Connector.SAPConnectorInfo]::get_Version()
        $PatchLevel = 
          [SAP.Middleware.Connector.SAPConnectorInfo]::get_KernelPatchLevel()
        $SAPRelease = 
          [SAP.Middleware.Connector.SAPConnectorInfo]::get_SAPRelease()

      Write-Host "`r`nNCo verion:" $Version
      Write-Host "Patch Level:" $PatchLevel
      Write-Host "SAP Release:" $SAPRelease

    }

  #-Sub Main------------------------------------------------------------
    Function Main () {
      If ($PSVersionTable.PSVersion.Major -ge 5) {
        Load-NCo
        Get-NCoVersion
      }
    }

  #-Main----------------------------------------------------------------
    Main

#-End-------------------------------------------------------------------
