Attribute VB_Name = "HelperListBox"
Option Explicit

Public Function cmdMoveUp_Click(lstMove As ListBox) As Integer
 'not by source
    Dim strTemp1 As String   '-- hold the selected index data temporarily for move
    
    Dim iCnt    As Integer   '-- holds the index of the item to be moved
    
    iCnt = lstMove.ListIndex
    
    If iCnt > -1 Then
         
         strTemp1 = lstMove.list(iCnt)
        
        '-- Add the item selected to one position above the current position
        lstMove.AddItem strTemp1, (iCnt - 1)
        
        '-- remove it from the current position. Note the current position has changed because the add has moved everything down by 1
        lstMove.RemoveItem (iCnt + 1)
        
        '-- Reselect the item that was moved.
             lstMove.Selected(iCnt - 1) = True
    
    End If
End Function

Public Function cmdMoveDown_Click(lstMove As ListBox) As Integer
    Dim strTemp1 As String    '-- hold the selected index data temporarily for move
    Dim iCnt    As Integer    '-- holds the index of the item to be moved
        
    '-- Assign the first index
    iCnt = lstMove.ListIndex
    
    If iCnt > -1 Then
         
         strTemp1 = lstMove.list(iCnt)
        
        '-- Add the item selected to below the current position
        lstMove.AddItem strTemp1, (iCnt + 2)
        
        lstMove.RemoveItem (iCnt)
        
        '-- Reselect the item that was moved.
        lstMove.Selected(iCnt + 1) = True
   End If

End Function
