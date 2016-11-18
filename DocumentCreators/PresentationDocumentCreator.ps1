
#-Begin-----------------------------------------------------------------
#-
#- Presentation document creator
#-
#- Author: Stefan Schnell
#-
#-----------------------------------------------------------------------

  #-Constants-----------------------------------------------------------
    $ppLayoutBlank = 12

    $ppSaveAsPresentation = 1               #PPT
    $ppSaveAsDefault = 11                   #PPTX
    $ppSaveAsGIF = 16
    $ppSaveAsJPG = 17
    $ppSaveAsPNG = 18
    $ppSaveAsBMP = 19
    $ppSaveAsTIF = 21
    $ppSaveAsEMF = 23
    $ppSaveAsPDF = 32                       #PDF
    $ppSaveAsOpenDocumentPresentation = 35  #ODP

    $msoTextOrientationHorizontal = 1
    $msoTextOrientationUpward = 2
    $msoTextOrientationDownward = 3
    $msoTextOrientationVertical = 5

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

  #-Sub Create-PresentationDocument-------------------------------------
  #-
  #- Creates different types of presentation documents automatically,
  #- with different count of slides and optional with a picture
  #-
  #---------------------------------------------------------------------
    Function Create-PresentationDocument($cntSlides, $PictureName, $Format) {

      If ($PictureName -eq $Null) {
        $WithPicture = $False
      }
      Else {
        $WithPicture = $True
        $PicName = $PSScriptRoot + "\" + $PictureName
      }

      $oPPT = New-Object -ComObject "PowerPoint.Application"
      $oPPT.Visible = [Int]$True
      $oPPT.Presentations.Add() > $Null

      For ($i = 1; $i -le $cntSlides; $i++) {
        $oSlide = $oPPT.ActivePresentation.Slides.Add(
          $oPPT.ActivePresentation.Slides.Count + 1, $ppLayoutBlank)

        $oTextBox = $oSlide.Shapes.AddTextBox($msoTextOrientationHorizontal, 
          10, 10, 480, 240)
        $oTextBox.TextFrame.TextRange.Text = $Text

        If ($WithPicture -eq $True) {
          $oSlide.Shapes.AddPicture($PicName, $True, $True, 480, 10, 200, 200)
        }

      }

      $GUID = [GUID]::NewGuid()
      $FileName = $env:USERPROFILE + "\Documents\" + $GUID
      $oPPT.ActivePresentation.SaveAs($FileName, $Format)
      $oPPT.ActivePresentation.Close()
      $oPPT.Quit()

    }

  #-Sub Main------------------------------------------------------------
    Function Main () {

      Create-PresentationDocument 10 "Gonzo.jpg" $ppSaveAsDefault

    }

  #-Main----------------------------------------------------------------
    If ($PSVersionTable.PSVersion.Major -ge 3) {
      Main
    }

#-End-------------------------------------------------------------------
