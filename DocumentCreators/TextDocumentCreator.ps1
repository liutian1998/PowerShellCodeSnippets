
#-Begin-----------------------------------------------------------------
#-
#- Text document creator
#-
#- Author: Stefan Schnell
#-
#-----------------------------------------------------------------------

  #-Constants-----------------------------------------------------------
    $wdLineBreak = 6
    $wdPageBreak = 7

    $wdStory = 6

    $wdFormatDocument = 0              #DOC
    $wdFormatRTF = 6                   #RTF
    $wdFormatDocumentDefault = 16      #DOCX
    $wdFormatPDF = 17                  #PDF
    $wdFormatOpenDocumentText = 23     #ODT

    $wdDoNotSaveChanges = 0

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
    $Text += " " + $Text

  #-Sub CreateTextDoc---------------------------------------------------
  #-
  #- Creates different types of text documents automatically, with
  #- different count of pages and optional with a picture
  #-
  #---------------------------------------------------------------------
    Function Create-TextDoc($cntPages, $PictureName, $Format) {

      If ($PictureName -eq $Null) {
        $WithPicture = $False
      }
      Else {
        $WithPicture = $True
        $PicName = $PSScriptRoot + "\" + $PictureName
      }

      $oWord = New-Object -ComObject "Word.Application"
      $oWord.Visible = $True

      $oWord.Documents.Add() > $Null
      For ($i = 1; $i -le $cntPages; $i++) {
        $oWord.Selection.TypeText($Text)
        If ($WithPicture -eq $True) {
          $oWord.Selection.InsertBreak($wdLineBreak)
          $oWord.Selection.EndKey($wdStory) > $Null
          $oWord.Selection.InlineShapes.AddPicture($PicName, $True, 
            $True) > $Null
        }
        If ($i -lt $cntPages) {
          $oWord.Selection.InsertBreak($wdPageBreak)
        }
      }

      $GUID = [GUID]::NewGuid()
      $FileName = $env:USERPROFILE + "\Documents\" + $GUID
      $oWord.ActiveDocument.SaveAs2([Ref]$FileName, 
        [Ref]$Format)
      $oWord.ActiveDocument.Close([Ref]$wdDoNotSaveChanges)
      $oWord.Quit()

    }

  #-Sub Main------------------------------------------------------------
    Function Main () {

      Create-TextDoc 10 "Gonzo.jpg" $wdFormatDocumentDefault

    }

  #-Main----------------------------------------------------------------
    If ($PSVersionTable.PSVersion.Major -ge 3) {
      Main
    }

#-End-------------------------------------------------------------------
