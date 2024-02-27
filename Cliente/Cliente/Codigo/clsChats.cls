VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const NUM_CHATS = 5

Dim MAX_CHARS_PER_LINE As Long

Private Type Chat
    active As Boolean
    life_time As Long
    start_time As Long
    text() As String
    font_index As Long
    font_color As Long
End Type

Dim chat_list(1 To NUM_CHATS) As Chat
Dim important_messages(1 To NUM_CHATS) As Chat

Dim chat_count As Long
Dim important_messages_count As Long

'Size of the screen to be used
Dim screen_width As Long
Dim screen_height As Long

'The engine, used to render text
Dim Engine As clsTileEngineX

'Gets number of ticks since windows started
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub Render()
    Dim cur_time As Long
    Dim LoopC As Long
    Dim offset As Long
    
    cur_time = GetTickCount
    
    offset = 16 * chat_count
    
    'Normal chats
    For LoopC = NUM_CHATS To 1 Step -1
        'Check if chat is empty
        If chat_list(LoopC).text(0) <> "" Then
            'Check if it´s alive
            If chat_list(LoopC).start_time + chat_list(LoopC).life_time < cur_time Or offset < 0 Then
                'It's dead
                ReDim chat_list(LoopC).text(0)
                chat_list(LoopC).active = False
                chat_count = chat_count - 1
            Else
                'Draw it
                Chat_Render chat_list(LoopC), offset
            End If
        End If
    Next LoopC
    
    offset = screen_height
    
    'Important messages
    For LoopC = NUM_CHATS To 1 Step -1
        'check if it´s empty
        If important_messages(LoopC).text(0) <> "" Then
            'Check if it´s alive
            If important_messages(LoopC).start_time + important_messages(LoopC).life_time < cur_time Then
                'It´s dead
                ReDim important_messages(LoopC).text(0)
                important_messages(LoopC).active = False
                important_messages_count = important_messages_count - 1
            Else
                'Draw it
                Chat_Render important_messages(LoopC), offset
            End If
        End If
    Next LoopC
End Sub

Private Sub Chat_Render(ByRef Chat As Chat, ByRef offset As Long)
    Dim LoopC As Long
    
    For LoopC = UBound(Chat.text()) To 0 Step -1
        'Print each line
        Engine.GUI_Text_Render Chat.text(LoopC), Chat.font_index, 0, offset, screen_width, 16, Chat.font_color, fa_left
        
        'Update offset
        offset = offset - 16
    Next LoopC
End Sub

Public Sub Chat_Add(ByVal text As String, ByVal font_index As Long, ByVal font_color As Long, Optional ByVal important_message As Boolean = False)
    Dim chat_index As Long
    Dim LoopC As Long
    
    chat_index = Chat_Next_Open(important_message)
    
    If important_message Then
        important_messages_count = important_messages_count + 1
        If important_messages_count > NUM_CHATS Then important_messages_count = NUM_CHATS
        
        If chat_index = NUM_CHATS Then
            For LoopC = 1 To 4
                important_messages(LoopC) = important_messages(LoopC + 1)
            Next LoopC
        End If
        
        'Important messages are displayed at the bottom
        Dialog_Format text, MAX_CHARS_PER_LINE, important_messages(chat_index).text()
        
        'Activate chat
        With important_messages(chat_index)
            .font_color = font_color
            .font_index = font_index
            .life_time = 15000
            .start_time = GetTickCount
            .active = True
        End With
    Else
        chat_count = chat_count + 1
        If chat_count > NUM_CHATS Then chat_count = NUM_CHATS
        
        'If we already completed the list rearrange it
        If chat_index = NUM_CHATS Then
            For LoopC = 1 To 4
                chat_list(LoopC) = chat_list(LoopC + 1)
            Next LoopC
        End If
        
        'Format text
        Dialog_Format text, MAX_CHARS_PER_LINE, chat_list(chat_index).text()
        
        'Activate chat
        With chat_list(chat_index)
            .font_color = font_color
            .font_index = font_index
            .life_time = 10000
            .start_time = GetTickCount
            .active = True
        End With
    End If
End Sub

Private Function Chat_Next_Open(ByVal important_message As Boolean)
    Dim LoopC As Long
    
    LoopC = 1
    
    If important_message Then
        Do Until important_messages(LoopC).active = False
            If LoopC = NUM_CHATS Then Exit Function
            
            LoopC = LoopC + 1
        Loop
    Else
        Do Until chat_list(LoopC).active = False
            If LoopC = NUM_CHATS Then Exit Function
            
            LoopC = LoopC + 1
        Loop
    End If
    
    Chat_Next_Open = LoopC
End Function

Public Function Initialize(ByRef tile_engine As clsTileEngineX, ByVal width_in_pixels As Long, ByVal height_in_pixels As Long) As Boolean
On Local Error GoTo ErrHandler
    Dim LoopC As Long
    
    'Set the engine object used to render chats
    If Not tile_engine Is Nothing Then Set Engine = tile_engine
    
    screen_height = height_in_pixels
    screen_width = width_in_pixels
    
    'Resize all chats to avoid run-time errors
    For LoopC = 1 To NUM_CHATS
        ReDim chat_list(LoopC).text(0)
        ReDim important_messages(LoopC).text(0)
    Next LoopC
    
    MAX_CHARS_PER_LINE = 120 * screen_width / 640
    
    Initialize = True
Exit Function

ErrHandler:
    MsgBox "An error ocurred when initializing the chat object."
End Function

Public Sub Engine_Resize(ByVal width_in_pixels As Long, ByVal height_in_pixels As Long)
    screen_height = height_in_pixels
    screen_width = width_in_pixels
    
    MAX_CHARS_PER_LINE = 120 * screen_width / 640
End Sub

Public Sub Dialog_Format(ByVal text As String, ByVal MAX_CHARS_PER_LINE As Long, ByRef formated_text() As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/27/2004
'Splits a message into several lines respecting the char per line limit
'**************************************************************
    Dim LoopC As Long
    Dim Line As Integer
    Dim temp_text() As String
    
    temp_text() = Split(text)
    
    ReDim formated_text(0)
    
    For LoopC = 0 To UBound(temp_text())
        'Check if next word fits into the current line
        If Len(formated_text(Line)) + Len(temp_text(LoopC)) < MAX_CHARS_PER_LINE Then
            'Avoid setting " " as the first char of every line
            If formated_text(Line) <> "" Then
                formated_text(Line) = formated_text(Line) & " " & temp_text(LoopC)
            Else
                formated_text(Line) = temp_text(LoopC)
            End If
        
        'Move on to next line
        Else
            Line = Line + 1
            ReDim Preserve formated_text(0 To Line)
            formated_text(Line) = temp_text(LoopC)
        End If
    Next LoopC
End Sub
