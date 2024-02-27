Attribute VB_Name = "HelperImage"
'This function returns image height, width and type
'of JPG, GIF, BMP & PNG formats.


'Type for returning image info
Public Type ImgDimType
  height As Long
  width As Long
End Type

'Inputs:
'
'fileName is a string containing the path name of the image file.
'
'ImgDim is passed as an empty type var and contains the height
'and width that's passed back.
'
'Ext is passed as an empty string and contains the image type
'as a 3 letter description that's passed back.
'
'
'Returns:
'
'True if the function was successful.
Function getImgDim(ByVal FileName As String, ImgDim As ImgDimType, _
                   ext As String) As Boolean



  'declare vars
  Dim handle As Integer, isValidImage As Boolean
  Dim byteArr(255) As Byte, i As Integer

  'init vars
  isValidImage = False
  ImgDim.height = 0
  ImgDim.width = 0
  
  'open file and get 256 byte chunk
  handle = FreeFile
  On Error GoTo endFunction
  Open FileName For Binary Access Read As #handle
  Get handle, , byteArr
  Close #handle
  

  'check for jpg header (SOI): &HFF and &HD8
  ' contained in first 2 bytes
  If byteArr(0) = &HFF And byteArr(1) = &HD8 Then
    isValidImage = True
  Else
    GoTo checkGIF
  End If
  
  'check for SOF marker: &HFF and &HC0 TO &HCF
  For i = 0 To 255
    If byteArr(i) = &HFF And byteArr(i + 1) >= &HC0 _
                         And byteArr(i + 1) <= &HCF Then
      ImgDim.height = byteArr(i + 5) * 256 + byteArr(i + 6)
      ImgDim.width = byteArr(i + 7) * 256 + byteArr(i + 8)
      Exit For
    End If
  Next i
  
  'get image type and exit
  ext = "jpg"
  GoTo endFunction


checkGIF:
  
  'check for GIF header
  If byteArr(0) = &H47 And byteArr(1) = &H49 And byteArr(2) = &H46 _
  And byteArr(3) = &H38 Then
    ImgDim.width = byteArr(7) * 256 + byteArr(6)
    ImgDim.height = byteArr(9) * 256 + byteArr(8)
    isValidImage = True
  Else
    GoTo checkBMP
  End If
  
  'get image type and exit
  ext = "gif"
  GoTo endFunction

  
checkBMP:
  
  'check for BMP header
  If byteArr(0) = 66 And byteArr(1) = 77 Then
    isValidImage = True
  Else
    GoTo checkPNG
  End If
  
  'get record type info
  If byteArr(14) = 40 Then
    
    'get width and height of BMP
    ImgDim.width = byteArr(21) * 256 ^ 3 + byteArr(20) * 256 ^ 2 _
                 + byteArr(19) * 256 + byteArr(18)
    
    ImgDim.height = byteArr(25) * 256 ^ 3 + byteArr(24) * 256 ^ 2 _
                  + byteArr(23) * 256 + byteArr(22)
  
  'another kind of BMP
  ElseIf byteArr(17) = 12 Then
  
    'get width and height of BMP
    ImgDim.width = byteArr(19) * 256 + byteArr(18)
    ImgDim.height = byteArr(21) * 256 + byteArr(20)
    
  End If
  
  'get image type and exit
  ext = "bmp"
  GoTo endFunction

  
checkPNG:

  'check for PNG header
  If byteArr(0) = &H89 And byteArr(1) = &H50 And byteArr(2) = &H4E _
  And byteArr(3) = &H47 Then
    ImgDim.width = byteArr(18) * 256 + byteArr(19)
    ImgDim.height = byteArr(22) * 256 + byteArr(23)
    isValidImage = True
  Else
    GoTo endFunction
  End If
  
  ext = "png"


endFunction:

  'return function's success status
  getImgDim = isValidImage
  

End Function



