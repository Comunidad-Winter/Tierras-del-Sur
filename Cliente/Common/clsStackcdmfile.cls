VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const BUFFER_SIZE As Long = 100

Private Items() As CDM_File

Private Index As Long

Public Property Get Count() As Long

   Count = Index

End Property

Public Sub Clear()

   ReDim Items(BUFFER_SIZE - 1)
   Index = 0

End Sub

Friend Sub Push(Item As CDM_File)

   If UBound(Items) = Index Then
      ReDim Preserve Items(Index + BUFFER_SIZE)
   End If


      Items(Index) = Item

   Index = Index + 1

End Sub

Friend Function Pop() As CDM_File

   Dim l As Long

   If Index > 0 Then
      Index = Index - 1

         Pop = Items(Index)


   End If

   If (UBound(Items) - Index) \ BUFFER_SIZE > 0 Then
      ReDim Preserve Items(UBound(Items) - BUFFER_SIZE)
   End If

End Function

Friend Function Peek() As CDM_File

   If Index > 0 Then
      Let Peek = Items(Index - 1)
   End If

End Function

Private Sub Class_Initialize()

   Clear

End Sub

Private Sub Class_Terminate()

   Erase Items

End Sub
