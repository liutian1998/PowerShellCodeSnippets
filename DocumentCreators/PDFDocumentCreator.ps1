
#-Begin-----------------------------------------------------------------
#-
#- PDF document creator
#-
#- Important hint: This script uses Debenu ActiveX Quick PDF Library
#- Lite, you can find it here:
#- http://www.debenu.com/products/development/debenu-pdf-library-lite/
#- You can find the forum for Quick PDF Library here:
#- http://www.quickpdf.org/forum/
#-
#- Important hint: To use the ActiveX library it is necessary to build
#- an interoperability library via type library importer TlbImp.exe.
#-
#- Author: Stefan Schnell
#-
#-----------------------------------------------------------------------

  #-Constants-----------------------------------------------------------
    $PDFMillimetres = 1
    $PDFInches = 2

    $PDFAlignmentCenter = 0
    $PDFAlignmentTop = 1
    $PDFAlignmentBottom = 2

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

  #-Sub CreatePDFDoc----------------------------------------------------
  #-
  #- Creates PDF documents automatically, with different count of
  #- pages and optional with a picture
  #-
  #---------------------------------------------------------------------
    Function Create-PDFDoc($cntPages, $PictureName) {

      If ($PictureName -eq $Null) {
        $WithPicture = $False
      }
      Else {
        $WithPicture = $True
        $PicName = $PSScriptRoot + "\" + $PictureName
      }

      $Path = $PSScriptRoot
      [Reflection.Assembly]::LoadFile($Path + 
        "\DebenuPDFLibraryLite1114.Interop.dll") > $Null

      $PDF = New-Object DebenuPDFLibraryLite1114.Interop.PDFLibraryClass
      $PDF.SetPageSize("A4") > $Null
      $PDF.SetMeasurementUnits($PDFMillimetres) > $Null

      For ($i = 1; $i -le $cntPages; $i++) {

        $PDF.DrawTextBox(10, 287, 190, 100, $Text, $PDFAlignmentTop) > $Null

        If ($WithPicture -eq $True) {
          $PDF.AddImageFromFile($PicName, 0) > $Null
          $ImgWidth = $PDF.ImageWidth() / 2
          $ImgHeight = $PDF.ImageHeight() / 2
          $PDF.DrawImage(10, 187, $ImgWidth, $ImgHeight) > $Null
        }

        If ($i -lt $cntPages) {
          $PDF.NewPage() > $Null
        }

      }

      $GUID = [GUID]::NewGuid()
      $FileName = $env:USERPROFILE + "\Documents\" + $GUID + ".pdf"
      $PDF.SaveToFile($FileName) > $Null

    }

  #-Sub Main------------------------------------------------------------
    Function Main() {

      Create-PDFDoc 10 "Gonzo.jpg"

    }

  #-Main----------------------------------------------------------------
    Main

#-End-------------------------------------------------------------------
