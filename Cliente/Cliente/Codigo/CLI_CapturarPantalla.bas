Attribute VB_Name = "CLI_CapturarPantalla"
Option Explicit

Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)


Public Sub CapturarPantalla()

  Dim FreeImage1 As Long
  Dim strFName As String
  Dim strFTemp As String

    If oJPG = 1 Then
        strFTemp = app.Path & "\fotos\screen.bmp"
        ' hide the form
        ' (as we don't want this in the screen shot)
        DoEvents
        Clipboard.Clear

        ' send a print screen button keypress event
        ' and DoEvents to allow windows time to process
        ' the event and capture the image to the clipboard
        keybd_event vbKeySnapshot, 0, 0, 0
        DoEvents
        ' send a print screen button up event
        keybd_event vbKeySnapshot, 0, &H2, 0
        DoEvents
        ' paste the clipboard contents into the picture box
        '[DEBUGUED BY WIZARD] Totalmente al pedo esto.
        'frmMain.ScreenCapture.Picture = Clipboard.GetData(vbCFBitmap)
        'DoEvents
        'DoEvents
        '[/WIZARD]
        ' change the pointer to an hourglass while the image is processed
        'save the image to a file using the application path
        ' [Wizard; Grabamos la imagen directamente desde el porta papeles]
        SavePicture Clipboard.GetData(vbCFBitmap), strFTemp
        DoEvents
        ' use the FreeImage.dll (http://freeimage.sourceforge.net/)
        ' to load the screen image
        FreeImage1 = FreeImage_Load(FIF_BMP, strFTemp, 0)
        ' save the screen capture as an JPEG image with high quality
        strFName = format(Now, "yyyy_mm_dd_hh_mm_ss")
        'strFName = Replace(Now, "/", "_")
        'strFName = Replace(strFName, ":", "_")
        'strFName = Replace(strFName, " ", "_")
        strFName = app.Path & "\fotos\TDS_foto_" & strFName & ".JPG"
        
        Call FreeImage_Save(FIF_jpeg, FreeImage1, strFName, &H80)
        'unload the images
        FreeImage_Unload (FreeImage1)
        ' restore the mouse pointer
        Kill strFTemp
      Else
        strFName = format(Now, "yyyy_mm_dd_hh_mm_ss")
        strFName = app.Path & "\fotos\TDS_foto_" & strFName & ".BMP"

        ' DoEvents
        Clipboard.Clear

        keybd_event vbKeySnapshot, 0, 0, 0
        ' DoEvents
        ' send a print screen button up event
        keybd_event vbKeySnapshot, 0, &H2, 0
        'DoEvents
        DoEvents
        SavePicture Clipboard.GetData(vbCFBitmap), strFName
        'DoEvents
    End If

    Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Guardado: " & strFName, 0, 200, 200, False, False, False)

End Sub
