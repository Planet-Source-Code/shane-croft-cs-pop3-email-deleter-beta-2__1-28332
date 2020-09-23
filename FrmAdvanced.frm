VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAdvanced 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Advanced Select"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAdvanced.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Option 3"
      Height          =   615
      Left            =   3000
      TabIndex        =   8
      Top             =   1200
      Width           =   3735
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         Max             =   1000
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Select"
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   240
         Left            =   2280
         TabIndex        =   11
         Text            =   "1"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   240
         Left            =   1320
         TabIndex        =   10
         Text            =   "1"
         Top             =   240
         Width           =   375
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         Max             =   1000
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         Caption         =   "to"
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Select Messages"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Option 2"
      Height          =   1215
      Left            =   3000
      TabIndex        =   4
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton Command1 
         Caption         =   "Select"
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "This will look in the ""From"", ""Subject"", and ""Date"""
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Select all that read "
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Option 1"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2775
      Begin VB.CommandButton Command4 
         Caption         =   "UnSelect All"
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Select All"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1320
      Picture         =   "FrmAdvanced.frx":0442
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2775
   End
End
Attribute VB_Name = "FrmAdvanced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Dim ItemFound As ListItem
Dim X As Long
X = 1

Do
Set ItemFound = FrmMain.ListView1.FindItem(Text1.Text, 1, X, lvwPartial)
ItemFound.EnsureVisible
ItemFound.Checked = True
ItemFound.Selected = True
X = X + 1
Loop Until X > FrmMain.ListView1.ListItems.Count
End Sub

Private Sub Command2_Click()
On Error Resume Next

If Text2.Text > Text3.Text Then
Exit Sub
End If

Dim Y As Long
Y = Text2.Text

Me.Enabled = False
Label1.Caption = "Please Wait..."
DoEvents
Dim Item As ListItem
Set Item = FrmMain.ListView1.ListItems(Y)
Set FrmMain.ListView1.SelectedItem = Item
DoEvents

Do Until Y = Text3.Text
FrmMain.ListView1.SelectedItem.Checked = True
DoEvents
FrmMain.ListView1.SelectedItem = FrmMain.ListView1.ListItems(FrmMain.ListView1.SelectedItem.Index + 1)
Y = Y + 1
Loop
FrmMain.ListView1.SelectedItem.Checked = True
DoEvents
Me.Enabled = True
Label1.Caption = "Done."
DoEvents
End Sub

Private Sub Command3_Click()
On Error Resume Next
Me.Enabled = False
Label1.Caption = "Please Wait..."
DoEvents
Dim Item As ListItem
Set Item = FrmMain.ListView1.ListItems(1)
Set FrmMain.ListView1.SelectedItem = Item
DoEvents

Do Until FrmMain.ListView1.SelectedItem.Index = FrmMain.ListView1.ListItems.Count
FrmMain.ListView1.SelectedItem.Checked = True
DoEvents
FrmMain.ListView1.SelectedItem = FrmMain.ListView1.ListItems(FrmMain.ListView1.SelectedItem.Index + 1)
Loop
FrmMain.ListView1.SelectedItem.Checked = True
DoEvents
Me.Enabled = True
Label1.Caption = "Done."
DoEvents
End Sub

Private Sub Command4_Click()
On Error Resume Next

Dim Item As ListItem
Set Item = FrmMain.ListView1.ListItems(1)
Set FrmMain.ListView1.SelectedItem = Item
DoEvents

Do Until FrmMain.ListView1.SelectedItem.Index = FrmMain.ListView1.ListItems.Count
FrmMain.ListView1.SelectedItem.Checked = False
DoEvents
FrmMain.ListView1.SelectedItem = FrmMain.ListView1.ListItems(FrmMain.ListView1.SelectedItem.Index + 1)
Loop
FrmMain.ListView1.SelectedItem.Checked = False
DoEvents
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command1_Click
 DoEvents
 End If
End Sub

Private Sub Text2_Change()
On Error Resume Next
UpDown1.Value = Text2.Text
DoEvents
End Sub

Private Sub Text3_Change()
On Error Resume Next
UpDown2.Value = Text3.Text
DoEvents
End Sub

Private Sub UpDown1_Change()
On Error Resume Next
Text2.Text = UpDown1.Value
End Sub

Private Sub UpDown2_Change()
On Error Resume Next
Text3.Text = UpDown2.Value
End Sub
