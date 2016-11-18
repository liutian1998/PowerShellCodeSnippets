
#-Begin-----------------------------------------------------------------
#-
#- Sheet document creator
#-
#- Author: Stefan Schnell
#-
#-----------------------------------------------------------------------

  #-Constants-----------------------------------------------------------
    $xlExcel9795 = 43                  #XLS
    $xlWorkbookDefault = 51            #XLSX
    $xlOpenDocumentSpreadsheet = 60    #ODS

  #-Variables-----------------------------------------------------------
    $Text  = "Lorem ipsum dolor sit amet, consetetur sadipscing elitr, "
    $Text += "sed diam nonumy eirmod tempor invidunt ut labore et dolor"
    $Text += "e magna aliquyam erat, sed diam voluptua. At vero eos et "
    $Text += "accusam et justo duo dolores et ea rebum. Stet clita kasd"
    $Text += " gubergren, no sea takimata sanctus est Lorem ipsum dolor"
    $Text += " sit amet. Lorem ipsum dolor sit amet, consetetur sadipsc"
    $Text += "ing elitr, sed diam nonumy eirmod tempor invidunt ut labo"
    $Text += "re et dolore magna aliquyam erat, sed diam voluptua. At v"
    $Text += "ero eos et accusam et justo duo dolores et ea rebum. Stet"
    $Text += " clita kasd gubergren, no sea takimata sanctus est Lorem "
    $Text += "ipsum dolor sit amet."

  #-Sub Create-SheetDocument--------------------------------------------
  #-
  #- Creates different types of table documents automatically with
  #- different count of sheets and optional with a picture
  #-
  #---------------------------------------------------------------------
    Function Create-SheetDocument($cntSheets, $PictureName, $Format) {

      If ($PictureName -eq $Null) {
        $WithPicture = $False
      }
      Else {
        $WithPicture = $True
        $PicName = $PSScriptRoot + "\" + $PictureName
      }

      $oExcel = New-Object -ComObject "Excel.Application"
      $oExcel.Visible = $True

      $oWorkBook = $oExcel.Workbooks.Add()
      For ($i = 1; $i -le $cntSheets; $i++) {
        Try {
          $oWorkBook.Sheets.Item($i).Select()
        }
        Catch {
          $oWorkBook.Sheets.Add([System.Reflection.Missing]::Value, 
            $oWorkBook.Sheets.Item($oWorkBook.Sheets.Count)) > $Null
        }

        For ($j = 1; $j -le 10; $j++) {
          $oWorkBook.Sheets.Item($i).Cells.Item($j, 1).Value2 = $Text
        }

        If ($WithPicture -eq $True) {
          $oWorkBook.Sheets.Item($i).Shapes.AddPicture($PicName, $True, 
            $True, 0, 200, 256, 256) > $Null
        }

      }

      $GUID = [GUID]::NewGuid()
      $FileName = $env:USERPROFILE + "\Documents\" + $GUID
      $oWorkBook.SaveAs($FileName, $Format)

      $oExcel.Quit()

    }


  #-Sub Main------------------------------------------------------------
    Function Main () {

      Create-SheetDocument 10 "Gonzo.jpg" $xlWorkbookDefault

    }

  #-Main----------------------------------------------------------------
    If ($PSVersionTable.PSVersion.Major -ge 3) {
      Main
    }

#-End-------------------------------------------------------------------
