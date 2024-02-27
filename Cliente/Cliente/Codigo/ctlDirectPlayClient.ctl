VERSION 5.00
Begin VB.UserControl ctlDirectPlayClient 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   PaletteMode     =   4  'None
   Picture         =   "ctlDirectPlayClient.ctx":0000
   ScaleHeight     =   555
   ScaleWidth      =   495
   Windowless      =   -1  'True
End
Attribute VB_Name = "ctlDirectPlayClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'ctlDirectPlayClient.ctl - ORE DirectPlay 8 Client - v0.5.0
'
'Handles all the TCP/IP traffic and ties all the client objects
'together.
'
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

'*****************************************************************
'Contributors History
'   When releasing modifications to this source file please add your
'   date of release, name, email, and any info to the top of this list.
'   Follow this template:
'    XX/XX/200X - Your Name Here (Your Email Here)
'       - Your Description Here
'       Sub Release Contributors:
'           XX/XX/2003 - Sub Contributor Name Here (SC Email Here)
'               - SC Description Here
'*****************************************************************
'
'Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com) - 8/20/2004
'   -Add: InventorySlotChanged event
'   -Add: Client_Request_Gump_Page and Receive_Gump methods
'   -Add: GumpLoaded Event
'   -Add: Chat_Type enumeration
'   -Add: Account_Char_Create method
'   -Add: Account_Char_Remove method
'   -Add: Client_Authenticate_With_Char method
'   -Add: Client_Info_Change method
'   -Add: Client_Logoff_Char method
'   -Add: Client_Logoff_Account method
'   -Add: Client_Logoff_Session_Terminate method
'   -Add: Client_Request_Char_Stats method
'   -Add: Client_Request_Stats_Roll method
'   -Add: NPC_Respond method
'   -Add: NPC_Talk method
'   -Add: Receive_NPC method
'   -Change: Receive_Chat now identifies all kinds of chat and displays it accordingly
'   -Change: private messages can now be sent
'   -Change: Receive event now identifies NPC server messages
'
'Aaron Perkins(aaron@baronsoft.com) - 8/04/2003
'   - First Release
'*****************************************************************

'***************************
'Required Externals
'***************************
'Reference to dx8vb.dll
'   - URL: http://www.microsoft.com/directx
'***************************
Option Explicit

'***************************
'Constants and enumerations
'***************************
Public Enum chat_type
    CT_Private
    CT_Normal
    CT_Map
    CT_Global
End Enum

'***************************
'Types
'***************************

'***************************
'Variables and objects
'***************************
Private DX As DirectX8                          'Main DirectX8 object
Private dp_client As DirectPlay8Client          'Server object, for message handling
Private dp_client_address As DirectPlay8Address 'Client's own IP, port
Private dp_server_address As DirectPlay8Address 'Server's own IP, port

Private Engine As clsTileEngineX
Private client_app_guid As String

Private client_char_id As Long

'Holds the index of the NPC speech (used to respond)
Private client_char_NPC_greet_index As Long

'***************************
'Arrays
'***************************

'***************************
'External Functions
'***************************
'Gets number of ticks since windows started
Private Declare Function GetTickCount Lib "kernel32" () As Long

'Very percise 64 bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

'***************************
'Events
'***************************
Public Event ClientConnected()
Public Event ClientConnectionFailed()
Public Event ClientDisconnected()
Public Event ClientAuthenticated(ByVal first_char_name As String, ByVal second_char_name As String, ByVal third_char_name As String, ByVal fourth_char_name As String)
Public Event EngineStop()
Public Event EngineStart()
Public Event ReceiveChatCritical(ByVal Chat_Text As String)
'Added by Juan Martín Sotuyo Dodero
Public Event StatsRolled(ByVal points As Long)
Public Event ReceiveCharStats(ByVal char_name As String, ByVal race As races, ByVal Class As classes, ByVal Alignment As Alignment, ByVal sphere As spheres, _
                                ByVal psionic_power As psionic_powers, ByVal level As Long, ByVal char_STR As Long, ByVal char_DEX As Long, ByVal char_CON As Long, _
                                ByVal char_INT As Long, ByVal char_WIS As Long, ByVal char_CHR As Long, ByVal portrait As Long, ByVal char_data_index As Long)
Public Event CharAuthenticated()
Public Event NPCChatReceive(ByVal NPC_greet As String, ByRef Responses() As String)
Public Event NPCChatStart(ByVal NPC_greet As String, ByRef Responses() As String, ByVal NPC_name As String, ByVal NPC_portrait As Long)
Public Event NPCChatTerminated()
Public Event ReceiveChatText(ByVal Chat_Text As String, ByVal Chat As chat_type, ByVal sender_id As Long)
Public Event GumpLoaded(ByVal cstring As String)
Public Event InventorySlotChanged(ByVal slot As Long, ByVal Item_Index As Long, ByVal Amount As Long, ByVal equiped As Boolean)

'***************************
Implements DirectPlay8Event
'***************************

Private Sub DirectPlay8Event_AddRemovePlayerGroup(ByVal message_id As Long, ByVal connection_id As Long, ByVal lGroupID As Long, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_AppDesc(byreffRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_AsyncOpComplete(byrefdpnotify As DxVBLibA.DPNMSG_ASYNC_OP_COMPLETE, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_ConnectComplete(ByRef dpnotify As DxVBLibA.DPNMSG_CONNECT_COMPLETE, ByRef fRejectMsg As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'
'**************************************************************
    If dpnotify.hResultCode = 0 Then
        RaiseEvent ClientConnected
    Else
        RaiseEvent ClientConnectionFailed
    End If
End Sub

Private Sub DirectPlay8Event_CreateGroup(ByVal group_id As Long, ByVal owner_id As Long, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_CreatePlayer(ByVal connection_id As Long, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_DestroyGroup(ByVal group_id As Long, ByVal reason_code As Long, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_DestroyPlayer(ByVal connection_id As Long, ByVal reason_code As Long, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_EnumHostsQuery(ByRef dpnotify As DxVBLibA.DPNMSG_ENUM_HOSTS_QUERY, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_EnumHostsResponse(ByRef dpnotify As DxVBLibA.DPNMSG_ENUM_HOSTS_RESPONSE, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_HostMigrate(ByVal new_host_id As Long, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_IndicateConnect(ByRef dpnotify As DxVBLibA.DPNMSG_INDICATE_CONNECT, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_IndicatedConnectAborted(ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_InfoNotify(ByVal message_id As Long, ByVal notify_id As Long, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_Receive(ByRef dpnotify As DxVBLibA.DPNMSG_RECEIVE, ByRef fRejectMsg As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/03/2004
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    'Get packet header
    Dim offset As Long
    
    'Get header
    Dim header As ServerPacketHeader
    Call GetDataFromBuffer(dpnotify.ReceivedData, header, SIZE_LONG, offset)
    
    'Get command
    Dim command As ServerPacketCommand
    Call GetDataFromBuffer(dpnotify.ReceivedData, command, SIZE_LONG, offset)
    
    'Get parameter(s)
    Dim received_data As String
    Dim parameters() As String
    Dim LoopC As Long
    Dim count As Long
    received_data = GetStringFromBuffer(dpnotify.ReceivedData, offset)
    count = General_Field_Count(received_data, P_DELIMITER_CODE)
    If count > 0 Then
        ReDim parameters(1 To count) As String
        For LoopC = 1 To count
            parameters(LoopC) = General_Field_Read(LoopC, received_data, P_DELIMITER_CODE)
        Next LoopC
    Else
        ReDim parameters(0 To 0) As String
    End If
    
    'Handle the packet
    If header = s_Player Then
        Receive_Player command, parameters()
        Exit Sub
    End If
    
    If header = s_Chat Then
        Receive_Chat command, parameters()
        Exit Sub
    End If
    
    If header = s_map Then
        Receive_Map command, parameters()
        Exit Sub
    End If
    
    If header = s_Char Then
        Receive_Char command, parameters()
        Exit Sub
    End If
    
    If header = s_NPC Then
        Receive_NPC command, parameters()
        Exit Sub
    End If
    
    If header = S_Gump Then
        Receive_Gump command, parameters()
        Exit Sub
    End If
End Sub

Private Sub DirectPlay8Event_SendComplete(ByRef dpnotify As DxVBLibA.DPNMSG_SEND_COMPLETE, ByRef fRejectMsg As Boolean)
    'We have to implement every method of this interface
End Sub

Private Sub DirectPlay8Event_TerminateSession(ByRef dpnotify As DxVBLibA.DPNMSG_TERMINATE_SESSION, ByRef fRejectMsg As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'
'**************************************************************
    RaiseEvent ClientDisconnected
End Sub

Private Sub UserControl_Initialize()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
End Sub

Private Sub UserControl_Terminate()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    Client_Deinitialize
End Sub

Public Function Client_Initialize(ByRef tile_engine As clsTileEngineX, ByVal app_guid As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'
'**************************************************************
'On Error GoTo Errorhandler
    Set DX = New DirectX8
    Set dp_client = DX.DirectPlayClientCreate
    Set dp_client_address = DX.DirectPlayAddressCreate
    Set dp_server_address = DX.DirectPlayAddressCreate
    dp_client.RegisterMessageHandler Me
    dp_client_address.SetSP DP8SP_TCPIP
    dp_server_address.SetSP DP8SP_TCPIP
    
    'Set app_guid
    client_app_guid = app_guid
    
    'Set pointer to engine
    Set Engine = tile_engine
    
   'Set delimiter
    P_DELIMITER = Chr$(P_DELIMITER_CODE)
    
    Client_Initialize = True
Exit Function
ErrorHandler:
    Client_Initialize = False
End Function

Public Sub Client_Deinitialize()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'
'**************************************************************
On Error Resume Next
    dp_client.CancelAsyncOperation 0, DPNCANCEL_ALL_OPERATIONS
    dp_client.Close
    dp_client.UnRegisterMessageHandler

    Set dp_client = Nothing
    Set dp_client_address = Nothing
    Set dp_server_address = Nothing
    Set DX = Nothing
End Sub

Public Function Client_Connect(ByVal server_ip As String, ByVal server_port As Long, ByVal account_name As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
On Error GoTo ErrorHandler:
    'Set account name in the client´s info
    Dim client_info As DPN_PLAYER_INFO
    client_info.name = account_name
    client_info.lInfoFlags = DPNINFO_NAME
    dp_client.SetClientInfo client_info, DPNOP_SYNC

    'Set server address
    dp_server_address.AddComponentString DPN_KEY_HOSTNAME, server_ip
    dp_server_address.AddComponentLong DPN_KEY_PORT, server_port

    'Connect
    Dim app_desc As DPN_APPLICATION_DESC
    app_desc.guidApplication = client_app_guid
    dp_client.Connect app_desc, dp_server_address, dp_client_address, 0, client_info.name, Len(client_info.name)
    
    Client_Connect = True
Exit Function
ErrorHandler:

End Function

Public Sub Client_Authenticate(ByVal password As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'
'**************************************************************
    'Send authentication string
    Send_Command c_Authenticate, c_Authenticate_Login, password
End Sub

Public Sub Client_Authenticate_New_Player(ByVal password As String, ByVal account_name As String, ByVal first_name As String, ByVal last_name As String, ByVal email As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/9/2003
'
'**************************************************************
    'Send authentication string
    Send_Command c_Authenticate, c_Authenticate_New, c_Packet_Player_New(password, account_name, first_name, last_name, email)
End Sub

Public Function Client_Connection_Status() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'
'**************************************************************
On Error GoTo ErrorHandler:
    If dp_client.GetServerInfo.name = "" Then
    End If
    Client_Connection_Status = True
Exit Function
ErrorHandler:
    Client_Connection_Status = False
End Function

Public Function Player_Move(ByVal Heading As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/17/2003
'
'**************************************************************
    'Find char_index based on server id
    Dim char_index As Long
    char_index = Engine.Char_Find(client_char_id)

    'Move the view position and the user_char
     If Engine.Map_Legal_Char_Pos_By_Heading(char_index, Heading) Then
         If Engine.Engine_View_Move(Heading) Then
            'Move player
             Engine.Char_Move char_index, Heading
            'Send Move command
            Send_Command c_Move, c_Move_Moved, CStr(Heading)
            Player_Move = True
         End If
     End If
End Function

Public Static Function Player_Attack() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'
'**************************************************************
    Dim frame_last_count As Long

    'Find char_index based on server id
    Dim char_index As Long
    char_index = Engine.Char_Find(client_char_id)

    'If it's been at least 16 frames since the last attack allow the command
    If Engine.Engine_Frame_Counter_Get - frame_last_count > 16 Then
        frame_last_count = Engine.Engine_Frame_Counter_Get
        Send_Command c_Action, c_Action_Attack, ""
        Player_Attack = True
    Else
        Player_Attack = False
    End If
End Function

Public Static Function Player_Item_Pickup() As Boolean
        Dim parameters As String
        Dim tmplong As Long
        Dim tmplong2 As Long
        Dim temp_x As Long
        Dim temp_y As Long
        
        Engine.Char_Map_Pos_Get Engine.Char_Find(client_char_id), temp_x, temp_y
        Engine.Map_Item_Get temp_x, temp_y, tmplong, tmplong
        If tmplong > 0 Then
            parameters = CStr(temp_x) _
                    & P_DELIMITER & CStr(temp_y)
            Send_Command c_Action, c_Action_Item_Pickup, parameters
            Player_Item_Pickup = True
            Exit Function
        End If
        Player_Item_Pickup = False
End Function

Private Sub Receive_Player(ByVal command As ServerPacketCommand, ByRef parameters() As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/13/2003
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
        'Player Authenticated
        If command = s_Player_Authenticated Then
            RaiseEvent ClientAuthenticated(parameters(1), parameters(2), parameters(3), parameters(4))
            Exit Sub
        End If
        
        'Engine Start
        If command = s_Player_Engine_Start Then
            RaiseEvent EngineStart
            Exit Sub
        End If
        
        'Engine Stop
        If command = s_Player_Engine_Stop Then
            RaiseEvent EngineStop
            Exit Sub
        End If
End Sub

Private Sub Receive_Gump(ByVal command As ServerPacketCommand, ByRef parameters() As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/03/2004
'
'**************************************************************
    If command = S_Gump_page Then
        RaiseEvent GumpLoaded(parameters(1))
        Exit Sub
    End If
End Sub

Private Sub Receive_Chat(ByVal command As ServerPacketCommand, ByRef parameters() As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/22/2004
'Modified by Juan Martín Sotuyo Dodero (Maraxus)
'**************************************************************
        'Whisper
        On Error Resume Next
        If command = s_Chat_Whisper Then
            RaiseEvent ReceiveChatText(parameters(1), CT_Private, 0)
            Exit Sub
        End If
        
        'Normal
        If command = s_Chat_Normal Then
            RaiseEvent ReceiveChatText(parameters(2), CT_Normal, CLng(parameters(1)))
            Exit Sub
        End If
        
        'Map
        If command = s_Chat_Map Then
            RaiseEvent ReceiveChatText(parameters(2), CT_Map, CLng(parameters(1)))
            Exit Sub
        End If
        
        'Global
        If command = s_Chat_Global Then
            RaiseEvent ReceiveChatText(parameters(1), CT_Global, 0)
            Exit Sub
        End If
        
        'Critical Chat
        If command = s_Chat_Critical Then
            RaiseEvent ReceiveChatCritical(parameters(1))
            Exit Sub
        End If
End Sub

Private Sub Receive_Map(ByVal command As ServerPacketCommand, ByRef parameters() As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/17/2003
'
'**************************************************************
        'Load Map
        If command = s_Map_Load Then
            Engine.Map_Load_Map_By_Name parameters(1)
            Exit Sub
        End If
        
        'Add item
        If command = s_Map_Item_Add Then
            Engine.Map_Item_Add CLng(parameters(1)), CLng(parameters(2)), CLng(parameters(4)), CLng(parameters(3))
            Engine.Map_Grh_Set CLng(parameters(1)), CLng(parameters(2)), CLng(parameters(5)), 5
            Exit Sub
        End If
        
        If command = s_Map_Item_Remove Then
            Engine.Map_Item_Remove CLng(parameters(1)), CLng(parameters(2))
            Engine.Map_Grh_UnSet CLng(parameters(1)), CLng(parameters(2)), 5
            Exit Sub
        End If
End Sub

Private Sub Receive_Char(ByVal command As ServerPacketCommand, ByRef parameters() As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/20/2004
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
        Dim char_index As Long

        'Stats rolled
        If command = s_Char_Stats_Rolled Then
            RaiseEvent StatsRolled(CLng(parameters(1)))
            Exit Sub
        End If
        
        'Char stats
        If command = s_Char_Stats_Get Then
            RaiseEvent ReceiveCharStats(parameters(1), CLng(parameters(2)), CLng(parameters(3)), CLng(parameters(4)), CLng(parameters(5)), CLng(parameters(6)), _
                                        CLng(parameters(7)), CLng(parameters(8)), CLng(parameters(9)), CLng(parameters(10)), CLng(parameters(11)), CLng(parameters(12)), _
                                        CLng(parameters(13)), CLng(parameters(14)), CLng(parameters(15)))
            Exit Sub
        End If
        
        'Char authenticated
        If command = s_Char_Authenticated Then
            RaiseEvent CharAuthenticated
        End If
        
        'Create Char
        If command = s_Char_Create Then
            Engine.Char_Create CLng(parameters(2)), CLng(parameters(3)), CLng(parameters(4)), _
                                CLng(parameters(5)), CLng(parameters(1))
            Exit Sub
        End If

       'Char ID
        If command = s_Char_ID_Set Then
            'Set player char id
            client_char_id = parameters(1)
            'Find char_index based on server id
            char_index = Engine.Char_Find(client_char_id)
            'Recenter view on player char
            Dim X As Long, Y As Long
            Engine.Char_Map_Pos_Get char_index, X, Y
            Engine.Engine_View_Pos_Set X, Y
            Exit Sub
        End If

        'Set Label
        If command = s_Char_Label_Set Then
            'Find char_index based on server id
            char_index = Engine.Char_Find(CLng(parameters(1)))
            'Set label
            Engine.Char_Label_Set char_index, parameters(2), 1
            Exit Sub
        End If
        
        'Set Char Data
        If command = s_Char_Data_Set Then
            'Find char_index based on server id
            char_index = Engine.Char_Find(CLng(parameters(1)))
            'Set
            Engine.Char_Data_Set char_index, CLng(parameters(2))
            Exit Sub
        End If
        
       'Set Char Body Data
        If command = s_Char_Data_Body_Set Then
            'Find char_index based on server id
            char_index = Engine.Char_Find(CLng(parameters(1)))
            'Set
            Engine.Char_Data_Body_Set char_index, CLng(parameters(2)), CByte(parameters(3))
            Exit Sub
        End If
        
        'Set Map Pos
        If command = s_Char_Pos_Set Then
            'Find char_index based on server id
            char_index = Engine.Char_Find(CLng(parameters(1)))
            'Set
            Engine.Char_Map_Pos_Set char_index, CLng(parameters(2)), CLng(parameters(3))
            'If client's char, center view on it
            If client_char_id = CLng(parameters(1)) Then
                Engine.Engine_View_Pos_Set CLng(parameters(2)), CLng(parameters(3))
            End If
            Exit Sub
        End If

        'Set Heading
        If command = s_Char_Heading_Set Then
            'Find char_index based on server id
            char_index = Engine.Char_Find(CLng(parameters(1)))
            'Set
            Engine.Char_Heading_Set char_index, CLng(parameters(2))
            Exit Sub
        End If

        'Move Char
        If command = s_Char_Move Then
            'Ignore client's own movment
            If client_char_id = CLng(parameters(1)) Then
                Exit Sub
            End If
            'Find char_index based on server id
            char_index = Engine.Char_Find(CLng(parameters(1)))
            'Move
            Engine.Char_Move char_index, CLng(parameters(2))
            Exit Sub
        End If
        
        'Remove Char
        If command = s_Char_Remove Then
            'Find char_index based on server id
            char_index = Engine.Char_Find(CLng(parameters(1)))
            'Remove
            Engine.Char_Remove char_index
            Exit Sub
        End If
        
        'Set Inventory Slot
        If command = s_Char_Set_Inventory_Slot Then
            RaiseEvent InventorySlotChanged(CLng(parameters(1)), CLng(parameters(2)), CLng(parameters(3)), CBool(parameters(4)))
        End If
End Sub

Public Sub Chat_Send(ByVal Message As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 3/24/2004
'Modified by Juan Martín Sotuyo Dodero (Maraxus)
'**************************************************************
    'TODO: More checks to make sure users aren't inputing garbage
     If Message <> "" Then
        'Send chat string
        If left(Message, 4) = "/to " Then
            Dim char_name As String
            char_name = General_Field_Read(2, Message, 127)
            Message = Right(Message, Len(Message) - 5 - Len(char_name))
            'Send_Command c_Chat, c_Chat_Whisper, Generic_Packet_Chat_Private_Message(char_name, message)
        ElseIf left(Message, 8) = "/scream " Then
            Message = Right(Message, Len(Message) - 8)
            Send_Command c_Chat, c_Chat_Map, Message
        Else
            Send_Command c_Chat, c_Chat_Normal, Message
        End If
    End If
End Sub

Private Sub Send_Command(ByVal header As ClientPacketHeader, ByVal command As ClientPacketCommand, ByRef parameters As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'
'**************************************************************
    'Don't send anything if not connected
    If Client_Connection_Status = False Then
        Exit Sub
    End If
        
    'New packet
    Dim byte_buffer() As Byte
    Dim offset As Long
    offset = NewBuffer(byte_buffer)
    
    'Add header
    Call AddDataToBuffer(byte_buffer, header, SIZE_LONG, offset)
    
    'Add command
    Call AddDataToBuffer(byte_buffer, command, SIZE_LONG, offset)

    'Add send_data
    Call AddStringToBuffer(byte_buffer, parameters, offset)

   'Send the packet, guarenteed, no loopback
    dp_client.Send byte_buffer, 0, DPNSEND_GUARANTEED Or DPNSEND_NOLOOPBACK
End Sub

Public Sub Client_Request_Stats_Roll()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/2/2004
'
'**************************************************************
    Send_Command c_Request, c_Request_Roll_Stats, ""
End Sub

Public Sub Client_Request_Char_Stats(ByVal char_name As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/4/2004
'
'**************************************************************
    Send_Command c_Request, c_Request_Char_Stats, char_name
End Sub

Public Sub Client_Logoff_Account()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/5/2004
'
'**************************************************************
    Send_Command c_Logoff, c_Logoff_Account, ""
End Sub

Public Sub Client_Info_Change(ByVal account_name As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/5/2004
'
'**************************************************************
    'Set account name in the client´s info
    Dim client_info As DPN_PLAYER_INFO
    client_info.name = account_name
    client_info.lInfoFlags = DPNINFO_NAME
    dp_client.SetClientInfo client_info, DPNOP_SYNC
End Sub

Public Sub Client_Account_Char_Remove(ByVal slot As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/5/2004
'
'**************************************************************
    Send_Command c_Account, c_Account_Remove_Char, CStr(slot)
End Sub

Public Sub Client_Account_Char_Create(ByVal char_name As String, ByVal race As races, ByVal Class As classes, _
                                        ByVal align As Alignment, ByVal sphere As spheres, ByVal psionic As psionic_powers, _
                                        ByVal char_STR As Long, ByVal char_DEX As Long, ByVal char_CON As Long, _
                                        ByVal char_INT As Long, ByVal char_WIS As Long, ByVal char_CHR As Long, _
                                        ByVal char_portrait As Long, ByVal char_data_index As Long, ByVal slot As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/5/2004
'
'**************************************************************
    Send_Command c_Account, c_Account_Add_Char, s_Packet_Character_Stats(char_name, race, Class, align, sphere, psionic, 0, char_STR, char_DEX, char_CON, char_INT, _
                char_WIS, char_CHR, char_portrait, char_data_index) & P_DELIMITER & CStr(slot)
End Sub

Public Sub Client_Logoff_Session_Terminate()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/5/2004
'
'**************************************************************
    Send_Command c_Logoff, c_Logoff_Session_Terminate, ""
End Sub

Public Sub Client_Authenticate_With_Char(ByVal char_name As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/7/2004
'
'**************************************************************
    Send_Command c_Authenticate, c_Authenticate_Char, char_name
End Sub

Public Sub Client_Logoff_Char()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/11/2004
'
'**************************************************************
    Send_Command c_Logoff, c_Logoff_Char, ""
End Sub

Private Sub Receive_NPC(ByVal command As ServerPacketCommand, ByRef parameters() As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/24/2004
'
'**************************************************************
    'Talk to NPC
    If command = s_NPC_Chat Then
        'Check if conversation ended
        If UBound(parameters()) = 0 Then
            RaiseEvent NPCChatTerminated
        Else
            Dim NumReplies As Long
            Dim Replies() As String
            Dim LoopC As Long
            
            'We retrieve the number of replies
            NumReplies = Val(parameters(2))
            
            'Prepare the reply array
            ReDim Replies(1 To NumReplies) As String
            For LoopC = 1 To NumReplies
                Replies(LoopC) = parameters(LoopC + 2)
            Next LoopC
            
            'We must check if it is a new NPC chat
            If UBound(parameters()) = NumReplies + 2 Then
                RaiseEvent NPCChatReceive(parameters(1), Replies())
                Exit Sub
            ElseIf UBound(parameters()) = NumReplies + 4 Then
                RaiseEvent NPCChatStart(parameters(1), Replies(), parameters(NumReplies + 3), CLng(parameters(NumReplies + 4)))
                Exit Sub
            End If
        End If
    End If
End Sub

Public Sub NPC_Talk(ByVal map_x As Long, ByVal map_y As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/22/2004
'
'**************************************************************
    Send_Command c_Action, c_Action_NPC_Chat, Generic_Packet_Map_Pos(map_x, map_y)
End Sub

Public Sub NPC_Respond(ByVal response_index As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/22/2004
'
'**************************************************************
    Send_Command c_Action, c_Action_NPC_Chat, CStr(response_index)
End Sub

Public Sub Client_Request_Gump_Page(ByVal id As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/03/2004
'
'**************************************************************
    Send_Command c_Gump, c_Gump_Page, CStr(id)
End Sub

Public Sub Client_Send_Gump_Button(ByVal params As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/03/2004
'
'**************************************************************
    Send_Command c_Gump, c_Gump_button, params
End Sub

Public Sub Client_Sysop_WorldSave()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    Send_Command c_Sysop, c_Sysop_Saveworld, ""
End Sub

Public Sub Client_Sysop_Reset()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    Send_Command c_Sysop, c_Sysop_Reset, ""
End Sub

Public Sub Client_Sysop_Shutdown()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    Send_Command c_Sysop, c_Sysop_Shutdown, ""
End Sub

Public Sub Client_Sysop_PlayerList()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    Send_Command c_Sysop, c_Sysop_Playerlist, ""
End Sub

Public Sub Client_Sysop_GoTo(ByVal Map As String, ByVal X As Long, ByVal Y As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    Send_Command c_Sysop, c_Sysop_Goto, Map & P_DELIMITER & Generic_Packet_Map_Pos(X, Y)
End Sub

Public Sub Client_Sysop_Summon(ByVal char_name As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    Send_Command c_Sysop, c_Sysop_Summon, char_name
End Sub

Public Sub Client_Sysop_Freeze()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    Send_Command c_Sysop, c_Sysop_Freeze, ""
End Sub

Public Sub Client_Sysop_Jail()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    Send_Command c_Sysop, c_Sysop_Jail, ""
End Sub

Public Sub Client_Sysop_Ban(ByVal ban As Boolean, ByVal char_name As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    Send_Command c_Sysop, c_Sysop_Ban, ban & P_DELIMITER & char_name
End Sub

Public Sub Client_Sysop_WorldMessage(ByVal Message As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    Send_Command c_Chat, c_Chat_Global, Message
End Sub

Public Sub Client_Sysop_Hide()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    Send_Command c_Sysop, c_Sysop_Hide, ""
End Sub

Public Sub Client_Sysop_Shide()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    Send_Command c_Sysop, c_Sysop_sHide, ""
End Sub

Public Sub Client_Sysop_Account()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    Send_Command c_Sysop, c_Sysop_Account, ""
End Sub

Public Sub Client_Sysop_Quest()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    Send_Command c_Sysop, c_Sysop_Quest, ""
End Sub

Public Sub Client_Sysop_NewItem(ByVal itemindex As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/05/2004
'
'**************************************************************
    Send_Command c_Sysop, c_Sysop_NewItem, CStr(itemindex)
End Sub
