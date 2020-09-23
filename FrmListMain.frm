VERSION 5.00
Begin VB.Form FrmListMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List Editor for CS POP3 Email Deleter"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmListMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   1635
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   4215
   End
   Begin VB.ListBox List1 
      Height          =   1635
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Double Click to remove from the list."
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Usernames"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Mail Servers"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "FrmListMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
Call List_Load(List1, App.Path & "\Servers.ini")
DoEvents
Call List_Load2(List2, App.Path & "\Usernames.ini")
DoEvents
End Sub
Public Sub List_Add(list As ListBox, txt As String)
On Error Resume Next
   List1.AddItem txt
End Sub

Public Sub List_Load(thelist As ListBox, FileName As String)
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
Public Sub List_Save(thelist As ListBox, FileName As String)
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
Public Sub List_Add2(list As ListBox, txt As String)
On Error Resume Next
   List2.AddItem txt
End Sub

Public Sub List_Load2(thelist As ListBox, FileName As String)
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
Public Sub List_Save2(thelist As ListBox, FileName As String)
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Call List_Save(List1, App.Path & "\Servers.ini")
DoEvents
Call List_Save2(List2, App.Path & "\Usernames.ini")
DoEvents
End Sub

Private Sub List1_DblClick()
On Error Resume Next
List1.RemoveItem List1.ListIndex
Err = 0
End Sub

Private Sub List2_DblClick()
On Error Resume Next
List2.RemoveItem List2.ListIndex
Err = 0
End Sub
