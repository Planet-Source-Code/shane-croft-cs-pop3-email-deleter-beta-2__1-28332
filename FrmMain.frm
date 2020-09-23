VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CS POP3 Email Deleter Beta 2"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   285
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   285
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3480
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Delete Mail"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9600
      Picture         =   "FrmMain.frx":0442
      TabIndex        =   14
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   3240
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Advanced Select"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8040
      TabIndex        =   11
      Top             =   5040
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4815
      Left            =   3600
      TabIndex        =   10
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   8493
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   1765
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "From"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Subject"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Size"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mail Server / Login"
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton Command2 
         Caption         =   "Connect"
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Clear"
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Hide Password"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   3255
      End
      Begin VB.ComboBox Combo2 
         Height          =   345
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1200
         Width           =   3255
      End
      Begin VB.ComboBox Combo1 
         Height          =   345
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Login Password"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Login Username"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Mail Server"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Mail Box Size:"
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   5400
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6960
      Picture         =   "FrmMain.frx":0884
      Top             =   5040
      Width           =   480
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Messages:"
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   5040
      Width           =   2895
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum POP3States
    POP3_Connect
    POP3_USER
    POP3_PASS
    POP3_STAT
    POP3_RETR
    POP3_DELE
    POP3_QUIT
End Enum

Private m_State         As POP3States

Private m_oMessage      As CMessage
Private m_colMessages   As New CMessages
Dim intIndex As Integer
Dim intStartOfString As Integer
Dim intEndOfString As Integer
Dim boolNotFound As Integer
Public Sub DisconnectMe()
On Error Resume Next
'm_State = POP3_QUIT
Winsock1.SendData "QUIT" & vbCrLf
Text2.Text = Text2.Text & "QUIT" & vbCrLf
Text2.SelStart = Len(Text2.Text)
Combo1.Enabled = True
Combo2.Enabled = True
Text1.Enabled = True
Check1.Enabled = True
Command1.Enabled = True
Command2.Caption = "Connect"
ListView1.ListItems.Clear
DoEvents
If ListView1.ListItems.Count = 0 Then
Command3.Enabled = False
Command5.Enabled = False
Else
Command3.Enabled = True
Command5.Enabled = True
End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
Text1.PasswordChar = "*"
Else
Text1.PasswordChar = ""
End If
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Combo2.SetFocus
 DoEvents
 End If
End Sub

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Text1.SetFocus
 DoEvents
 End If
End Sub

Private Sub Command1_Click()

Combo1.Text = ""
Combo2.Text = ""
Text1.Text = ""

End Sub

Public Sub Command2_Click()
'On Error Resume Next
Dim YY As String
Dim ZZ As String

If Command2.Caption = "Disconnect" Then
Call DisconnectMe
Exit Sub
End If

If Combo1.Text = "" Then
MsgBox "Please Enter A Mail Server!"
Combo1.SetFocus
Exit Sub
Else
YY = Combo1.Text
List1.AddItem Combo1.Text
DoEvents
Call xListKillDupes(List1)
DoEvents
Call List_Save(List1, App.Path & "\Servers.ini")
DoEvents
Combo1.Clear
List1.Clear
DoEvents
Combo1.Text = YY
Call List_Load(List1, App.Path & "\Servers.ini")
DoEvents
Call Combo_Load(Combo1, App.Path & "\Servers.ini")
DoEvents
End If

If Combo2.Text = "" Then
MsgBox "Please Enter A User Name!"
Combo2.SetFocus
Exit Sub
Else
ZZ = Combo2.Text
List2.AddItem Combo2.Text
DoEvents
Call xListKillDupes(List2)
DoEvents
Call List_Save2(List2, App.Path & "\Usernames.ini")
DoEvents
Combo2.Clear
List2.Clear
DoEvents
Combo2.Text = ZZ
Call List_Load2(List2, App.Path & "\Usernames.ini")
DoEvents
Call Combo_Load2(Combo2, App.Path & "\Usernames.ini")
DoEvents
End If

If Text1.Text = "" Then
MsgBox "Please Enter A Password!"
Text1.SetFocus
Exit Sub
End If

Combo1.Enabled = False
Combo2.Enabled = False
Text1.Enabled = False
Check1.Enabled = False
Command1.Enabled = False

Command2.Caption = "Disconnect"
ListView1.ListItems.Clear
Text2.Text = ""
m_colMessages.Clear
DoEvents
    'Check the emptiness of all the text fields except for the txtBody
    '
    'Change the value of current session state
    m_State = POP3_Connect
    '
    'Close the socket in case it was opened while another session
    Winsock1.Close
    '
    'reset the value of the local port in order to let to the
    'Windows Sockets select the new one itself
    'It's necessary in order to prevent the "Address in use" error,
    'which can appear if the Winsock Control has already used while the 
    'previous session
    Winsock1.LocalPort = 0
    '
    'POP3 server waits for the connection request at the port 110.
    'According with that we want the Winsock Control to be connected to
    'the port number 110 of the server we have supplied in combo1 field
    Winsock1.Connect Combo1, 110
End Sub

Private Sub Command3_Click()
FrmAdvanced.Show
DoEvents
FrmAdvanced.UpDown1.Max = ListView1.ListItems.Count
FrmAdvanced.UpDown2.Max = ListView1.ListItems.Count
FrmAdvanced.Text3.Text = ListView1.ListItems.Count
End Sub

Private Sub Command5_Click()
Dim Item As ListItem
Set Item = ListView1.ListItems(1)
Set ListView1.SelectedItem = Item
DoEvents

Do Until ListView1.SelectedItem.Index = ListView1.ListItems.Count
If ListView1.SelectedItem.Checked = True Then
ListView1.SelectedItem.EnsureVisible
Text2.Text = Text2.Text & "DELE " & ListView1.SelectedItem.Text & vbCrLf
Text2.SelStart = Len(Text2.Text)
DoEvents
m_State = POP3_DELE
Winsock1.SendData "DELE " & ListView1.SelectedItem.Text & vbCrLf
DoEvents
ListView1.SelectedItem.Checked = False
End If
ListView1.SelectedItem = ListView1.ListItems(ListView1.SelectedItem.Index + 1)
Loop
If ListView1.SelectedItem.Checked = True Then
ListView1.SelectedItem.EnsureVisible
Text2.Text = Text2.Text & "DELE " & ListView1.SelectedItem.Text & vbCrLf
Text2.SelStart = Len(Text2.Text)
DoEvents
m_State = POP3_DELE
Winsock1.SendData "DELE " & ListView1.SelectedItem.Text & vbCrLf
DoEvents
ListView1.SelectedItem.Checked = False
End If
End Sub
Private Sub Form_Load()
Dim fFile As Integer
Dim XXXX As String
fFile = FreeFile

Open App.Path & "\Settings.ini" For Input As fFile
Input #fFile, XXXX
Close fFile
Check1.Value = XXXX
DoEvents

Call List_Load(List1, App.Path & "\Servers.ini")
DoEvents
Call Combo_Load(Combo1, App.Path & "\Servers.ini")
DoEvents
Call List_Load2(List2, App.Path & "\Usernames.ini")
DoEvents
Call Combo_Load2(Combo2, App.Path & "\Usernames.ini")
DoEvents
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim fFile As Integer
fFile = FreeFile

Open App.Path & "\Settings.ini" For Output As fFile
Print #fFile, Check1.Value
Close fFile
DoEvents

End
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command2_Click
 DoEvents
 End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
    Dim strData As String
    
    Static intMessages          As Integer 'the number of messages to be loaded
    Static intCurrentMessage    As Integer 'the counter of loaded messages
    Static strBuffer            As String  'the buffer of the loading message
    Static TotalSize            As Long
    Dim EmailNum                As Long
    '
    'Save the received data into strData variable
    Winsock1.GetData strData
    'Text2.Text = Text2.Text & strData & vbCrLf
    Text2.SelStart = Len(Text2.Text)
    
    If Left$(strData, 1) = "+" Or m_State = POP3_RETR Then
        'If the first character of the server's response is "+" then
        'server accepted the client's command and waits for the next one
        'If this symbol is "-" then here we can do nothing
        'and execution skips to the Else section of the code
        'The first symbol may differ from "+" or "-" if the received
        'data are the part of the message's body, i.e. when
        'm_State = POP3_RETR (the loading of the message state)
        Select Case m_State
            Case POP3_Connect
                '
                'Reset the number of messages
                intMessages = 0
                intCurrentMessage = 0
                '
                'Change current state of session
                m_State = POP3_USER
                '
                'Send to the server the USER command with the parameter.
                'The parameter is the name of the mail box
                'Don't forget to add vbCrLf at the end of the each command!
                Winsock1.SendData "USER " & Combo2 & vbCrLf
                Text2.Text = Text2.Text & "USER " & Combo2 & vbCrLf
                Text2.SelStart = Len(Text2.Text)
                'Here is the end of Winsock1_DataArrival routine until the
                'next appearing of the DataArrival event. But next time this
                'section will be skipped and execution will start right after
                'the Case POP3_USER section.
            Case POP3_USER
                '
                'This part of the code runs in case of successful response to
                'the USER command.
                'Now we have to send to the server the user's password
                '
                'Change the state of the session
                m_State = POP3_PASS
                Winsock1.SendData "PASS " & Text1 & vbCrLf
                Text2.Text = Text2.Text & "PASS ***** " & vbCrLf
                Text2.SelStart = Len(Text2.Text)
            Case POP3_PASS
                '
                'The server answered positively to the process of the
                'identification and now we can send the STAT command. As a
                'response the server is going to return the number of
                'messages in the mail box and its size in octets
                '
                ' Change the state of the session
                m_State = POP3_STAT
                '
                'Send STAT command to know how many
                'messages in the mailbox
                Winsock1.SendData "STAT" & vbCrLf
                Text2.Text = Text2.Text & "STAT" & vbCrLf
                Text2.SelStart = Len(Text2.Text)
            Case POP3_STAT
                '
                'The server's response to the STAT command looks like this:
                '"+OK 0 0" (no messages at the mailbox) or "+OK 3 7564"
                '(there are messages). Evidently, the first of all we have to
                'find out the first numeric value that contains in the
                'server's response
                intMessages = Get_After_Seperator(strData, 1, " ")
                TotalSize = Get_After_Seperator(strData, 2, " ")
                Label4.Caption = "Messages: " & intMessages
                Label5.Caption = "Mail Box Size: " & Format(TotalSize / 1024, 0) & " KB"
                DoEvents
                If intMessages = 0 Then
                MsgBox "There are no messages on the server."
                Winsock1.SendData "QUIT" & vbCrLf
                Text2.Text = Text2.Text & "QUIT" & vbCrLf
                Text2.SelStart = Len(Text2.Text)
                Winsock1.Close
                Call DisconnectMe
                Exit Sub
                End If
        If MsgBox("There Is " & intMessages & " Messages!" & vbCrLf & "Do you wish to list all emails?", vbYesNo, "Email Found") = vbNo Then
        'msgbox asking if user wants to goto www.mail.com to read new mail
        If MsgBox("Would you like to delete all the messages on the server?" & vbCrLf & "Saying yes will delete all mail on the server." & vbCrLf & "Saying no will keep all mail on the server and disconnect.", vbYesNo, "Email Found") = vbNo Then
        Winsock1.SendData "QUIT" & vbCrLf
        Text2.Text = Text2.Text & "QUIT" & vbCrLf
        Text2.SelStart = Len(Text2.Text)
        Winsock1.Close
        Call DisconnectMe
        Exit Sub
        Else
        EmailNum = 1
        Do Until EmailNum = intMessages + 1
        Text2.Text = Text2.Text & "DELE " & EmailNum & vbCrLf
        Text2.SelStart = Len(Text2.Text)
        DoEvents
        m_State = POP3_DELE
        Winsock1.SendData "DELE " & EmailNum & vbCrLf
        DoEvents
        EmailNum = EmailNum + 1
        Loop
        Call DisconnectMe
        Exit Sub
        End If
        End If
                If intMessages > 0 Then
                    '
                    'Oops. There is something in the mailbox!
                    'Change the session state
                    m_State = POP3_RETR
                    '
                    'Increment the number of messages by one
                    intCurrentMessage = intCurrentMessage + 1
                    '
                    'and we're sending to the server the RETR command in
                    'order to retrieve the first message
                    Winsock1.SendData "RETR 1" & vbCrLf
                    Text2.Text = Text2.Text & "RETR 1" & vbCrLf
                    Text2.SelStart = Len(Text2.Text)
                Else
                    'The mailbox is empty. Send the QUIT command to the
                    'server in order to close the session
                    m_State = POP3_QUIT
                    Winsock1.SendData "QUIT" & vbCrLf
                    Text2.Text = Text2.Text & "QUIT" & vbCrLf
                    Text2.SelStart = Len(Text2.Text)
                    MsgBox "You have not mail.", vbInformation
                End If
            Case POP3_RETR
                'This code executes while the retrieving of the mail body
                'The size of the message could be quite big and the
                'DataArrival event may rise several time. All the received
                'data stores at the strBuffer variable:
                strBuffer = strBuffer & strData
                '
                'If case of presence of the point in the buffer it indicates
                'the end of the message (look at SMTP protocol)
                If InStr(1, strBuffer, vbLf & "." & vbCrLf) Then
                    '
                    'Done! The message has loaded
                    '
                    'Delete the first string-the server's response
                    strBuffer = Mid$(strBuffer, InStr(1, strBuffer, vbCrLf) + 2)
                    '
                    'Delete the last string. It contains only the "." symbol,
                    'which indicates the end of the message
                    strBuffer = Left$(strBuffer, Len(strBuffer) - 3)
                    '
                    'Add new message to m_colMessages collection
                    Set m_oMessage = New CMessage
                    m_oMessage.CreateFromText strBuffer
                    m_colMessages.Add m_oMessage, CStr(intCurrentMessage)
                   
Dim lvItem As ListItem


        Set lvItem = ListView1.ListItems.Add(, , CStr(intCurrentMessage))
        lvItem.SubItems(1) = m_oMessage.From
        lvItem.SubItems(2) = m_oMessage.Subject
        lvItem.SubItems(3) = m_oMessage.SendDate
        lvItem.SubItems(4) = m_oMessage.Size
        lvItem.EnsureVisible

                     Set m_oMessage = Nothing
                     m_colMessages.Clear
                     DoEvents
                    'Clear buffer for next message
                    strBuffer = ""
                    'Now we comparing the number of loaded messages with the
                    'one returned as a response to the STAT command
                    If intCurrentMessage = intMessages Then
                        'If these values are equal then all the messages
                        'have loaded. Now we can finish the session. Due to
                        'this reason we send the QUIT command to the server
                    'm_State = POP3_QUIT
                    'Winsock1.SendData "QUIT" & vbCrLf
                    'Text2.Text = Text2.Text & "QUIT" & vbCrLf
                    If ListView1.ListItems.Count = 0 Then
                    Command3.Enabled = False
                    Command5.Enabled = False
                    Else
                    Command3.Enabled = True
                    Command5.Enabled = True
                    End If
                    Else
                        'If these values aren't equal then there are
                        'remain messages. According with that
                        'we increment the messages' counter
                        intCurrentMessage = intCurrentMessage + 1
                        '
                        'Change current state of session
                        m_State = POP3_RETR
                        '
                        'Send RETR command to download next message
                        Winsock1.SendData "RETR " & _
                        CStr(intCurrentMessage) & vbCrLf
                        Text2.Text = Text2.Text & "RETR " & intCurrentMessage & vbCrLf
                        Text2.SelStart = Len(Text2.Text)
                    End If
                End If
            Case POP3_DELE
            m_State = POP3_DELE
            'Winsock1.SendData "DELE 1" & vbCrLf
            
            Case POP3_QUIT
                'No matter what data we've received it's important
                'to close the connection with the mail server
                Winsock1.Close
                'Now we're calling the ListMessages routine in order to
                'fill out the ListView control with the messages we've          
                'downloaded
                'Call ListMessages
        End Select
    Else
        'As you see, there is no sophisticated error
        'handling. We just close the socket and show the server's response
        'That's all. By the way even fully featured mail applications
        'do the same.
            Winsock1.Close
            MsgBox "POP3 Error: " & strData, _
            vbExclamation, "POP3 Error"
    End If
End Sub
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim ErrorStats As String
Winsock1.Close
    ErrorStats = Number & " : " & Description
    Text2.Text = Text2.Text & ErrorStats & vbCrLf
    Text2.SelStart = Len(Text2.Text)
End Sub

'Private Sub ListMessages()
'On Error Resume Next
'Dim oMes As CMessage
'Dim lvItem As ListItem
'
'    For Each oMes In m_colMessages
'        Set lvItem = ListView1.ListItems.Add(, , oMes.MailNumber)
'        lvItem.Key = oMes.MessageID
'        lvItem.SubItems(1) = oMes.From
'        lvItem.SubItems(2) = oMes.Subject
'        lvItem.SubItems(3) = oMes.SendDate
'        lvItem.SubItems(4) = oMes.Size
'    Next
'
'End Sub

Function Get_After_Seperator(ByVal strString As String, ByVal intNthOccurance As Integer, ByVal strSeperator As String) As String
    'On Error Resume Next

    
    'check for intNthOccurance = 0--i.e. fir
    '     st one


    If (intNthOccurance = 0) Then


        If (InStr(strString, strSeperator) > 0) Then
                Get_After_Seperator = Left(strString, InStr(strString, strSeperator) - 1)
        Else
                Get_After_Seperator = strString
        End If
    Else
        'not the first one
        'init start of string on first comma
        intStartOfString = InStr(strString, strSeperator)
        
        'place start of string after intNthOccur
        '     ance-th comma (-1 since
        'already did one
        boolNotFound = 0


        For intIndex = 1 To intNthOccurance - 1
            'get next comma
            intStartOfString = InStr(intStartOfString + 1, strString, strSeperator)
            'check for not found


            If (intStartOfString = 0) Then
                boolNotFound = 1
            End If
        Next intIndex
        
        'put start of string past 1st comma
        intStartOfString = intStartOfString + 1
        
        'check for ending in a comma


        If (intStartOfString > Len(strString)) Then
            boolNotFound = 1
        End If
        


        If (boolNotFound = 1) Then
            Get_After_Seperator = "NOT FOUND"
        Else
            intEndOfString = InStr(intStartOfString, strString, strSeperator)
            
            ' check for no second comma (i.e. end of
            '     string)


            If (intEndOfString = 0) Then
                intEndOfString = Len(strString) + 1
            Else
                intEndOfString = intEndOfString - 1
            End If
            Get_After_Seperator = Mid$(strString, intStartOfString, intEndOfString - intStartOfString + 1)
        End If
    End If
End Function
Public Sub List_Add(list As listbox, txt As String)
On Error Resume Next
   List1.AddItem txt
End Sub

Public Sub List_Load(thelist As listbox, FileName As String)
    'Loads a file to a list box
    On Error Resume Next
    Dim TheContents As String
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Input As fFile
    Do
        Line Input #fFile, TheContents$
        If TheContents$ = "" Then
        Else
        Call List_Add(List1, TheContents$)
        End If
    Loop Until EOF(fFile)
    Close fFile
End Sub
Public Sub Combo_Add(list As ComboBox, txt As String)
On Error Resume Next
   Combo1.AddItem txt
End Sub

Public Sub Combo_Load(thelist As ComboBox, FileName As String)
    'Loads a file to a list box
    On Error Resume Next
    Dim TheContents As String
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Input As fFile
    Do
        Line Input #fFile, TheContents$
        If TheContents$ = "" Then
        Else
        Call Combo_Add(Combo1, TheContents$)
        End If
    Loop Until EOF(fFile)
    Close fFile
End Sub
Public Sub List_Save(thelist As listbox, FileName As String)
    'Save a listbox as FileName
    On Error Resume Next
    Dim Save As Long
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Output As fFile
    For Save = 0 To thelist.ListCount - 1
        Print #fFile, List1.list(Save)
    Next Save
    Close fFile
End Sub
Public Sub List_Add2(list As listbox, txt As String)
On Error Resume Next
   List2.AddItem txt
End Sub

Public Sub List_Load2(thelist As listbox, FileName As String)
    'Loads a file to a list box
    On Error Resume Next
    Dim TheContents As String
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Input As fFile
    Do
        Line Input #fFile, TheContents$
        If TheContents$ = "" Then
        Else
        Call List_Add2(List2, TheContents$)
        End If
    Loop Until EOF(fFile)
    Close fFile
End Sub
Public Sub Combo_Add2(list As ComboBox, txt As String)
On Error Resume Next
   Combo2.AddItem txt
End Sub

Public Sub Combo_Load2(thelist As ComboBox, FileName As String)
    'Loads a file to a list box
    On Error Resume Next
    Dim TheContents As String
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Input As fFile
    Do
        Line Input #fFile, TheContents$
        If TheContents$ = "" Then
        Else
        Call Combo_Add2(Combo2, TheContents$)
        End If
    Loop Until EOF(fFile)
    Close fFile
End Sub
Public Sub List_Save2(thelist As listbox, FileName As String)
    'Save a listbox as FileName
    On Error Resume Next
    Dim Save As Long
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Output As fFile
    For Save = 0 To thelist.ListCount - 1
        Print #fFile, List2.list(Save)
    Next Save
    Close fFile
End Sub

