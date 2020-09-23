VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{92DC12AE-4D3B-4731-A352-709AA2D97A48}#1.0#0"; "DGSSPLITBAR.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   Caption         =   "Splitbar Example"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   522
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   626
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   2055
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "Form1.frx":0000
      Top             =   4200
      Width           =   2535
   End
   Begin DGSsplitBar.splitbar splitbar2 
      Height          =   75
      Left            =   0
      TabIndex        =   12
      Top             =   4035
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   132
      BackColor       =   -2147483645
   End
   Begin DGSsplitBar.splitbar splitbar1 
      Height          =   285
      Left            =   0
      TabIndex        =   8
      Top             =   2640
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   503
      BackColor       =   -2147483644
      MinTopHeight    =   15
      Begin VB.TextBox Text3 
         Height          =   240
         Left            =   690
         TabIndex        =   17
         Text            =   "This splitbar contains other controls"
         Top             =   15
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Height          =   255
         Left            =   360
         MousePointer    =   1  'Arrow
         Picture         =   "Form1.frx":0087
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   270
      End
      Begin VB.CommandButton Command3 
         Height          =   255
         Left            =   0
         MousePointer    =   1  'Arrow
         Picture         =   "Form1.frx":05C9
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   285
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   255
         Left            =   6120
         MousePointer    =   1  'Arrow
         TabIndex        =   9
         Top             =   15
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   1080
      Left            =   0
      ScaleHeight     =   1020
      ScaleWidth      =   7515
      TabIndex        =   5
      Top             =   2940
      Width           =   7575
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   0
         Picture         =   "Form1.frx":0B0B
         ScaleHeight     =   1005
         ScaleWidth      =   1905
         TabIndex        =   6
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Again this picture control is added to the splitbar, but no code has been added to the picture resize event"
         Height          =   855
         Left            =   4200
         TabIndex        =   14
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000008&
         Caption         =   "Label2"
         Height          =   1035
         Left            =   2160
         TabIndex        =   7
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H8000000B&
      Height          =   2475
      Left            =   0
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   501
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.TextBox Text1 
         Height          =   1875
         Left            =   15
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "Form1.frx":12AE7
         Top             =   240
         Width           =   3120
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   660
         Left            =   3360
         Picture         =   "Form1.frx":12BDE
         ScaleHeight     =   630
         ScaleWidth      =   1185
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   330
         Left            =   5940
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1275
         Left            =   4680
         TabIndex        =   2
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2249
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Test"
            Object.Width           =   38100
         EndProperty
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "The mintopheight has been set so that this label is always visible."
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   7575
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   $"Form1.frx":24BBA
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   2640
      TabIndex        =   16
      Top             =   4200
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
splitbar1.Top = Picture4.Top + splitbar1.MinTopHeight
splitbar1.Update
splitbar1.SetFocus
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 0
End Sub

Private Sub Command4_Click()
splitbar1.Top = (Picture5.Top + Picture5.Height) - 30
splitbar1.Update
splitbar1.SetFocus
End Sub

Private Sub Form_Load()
splitbar1.AddControlTop Picture4
splitbar1.AddControlBottom Picture5
splitbar1.Update

splitbar2.AddControlTop Picture5
splitbar2.AddControlBottom Label1
splitbar2.AddControlBottom Text2
splitbar2.Update







ListView1.ListItems.Add , , "some test text"
End Sub


Private Sub Picture3_Resize()
Text2.Top = Picture3.ScaleTop
Text2.Height = Picture3.ScaleHeight
End Sub

Private Sub Picture4_Resize()
On Error Resume Next
Text1.Height = (Picture4.ScaleHeight - Text1.Top)
Picture1.Height = (Picture4.ScaleHeight - 5)
ListView1.Height = (Picture4.ScaleHeight - 5)
End Sub

