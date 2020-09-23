VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "Picclp32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   855
   ClientLeft      =   3300
   ClientTop       =   345
   ClientWidth     =   3615
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Image1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   200
      ScaleHeight     =   135
      ScaleWidth      =   3255
      TabIndex        =   2
      ToolTipText     =   "Click to toggle, Right click to exit."
      Top             =   720
      Width           =   3255
   End
   Begin PicClip.PictureClip PictureClip1 
      Left            =   120
      Top             =   120
      _ExtentX        =   5741
      _ExtentY        =   529
      _Version        =   393216
      Rows            =   2
      Picture         =   "Form1.frx":030A
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":364C
      Left            =   360
      List            =   "Form1.frx":3659
      TabIndex        =   1
      Text            =   "Small Icons"
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    TextBackGroundTransparent True
    Select Case Combo1.Text
        Case "Large Icons"
            ChangeExplorerListViewStyle LVS_ICON
        Case "Small Icons"
            ChangeExplorerListViewStyle LVS_SMALLICON
        Case "List View"
            ChangeExplorerListViewStyle LVS_LIST
    End Select
End Sub
Private Sub Form_Load()
    Me.Move (Screen.Width / 2) - (Me.Width / 2), -Image1.Top
    Image1.Picture = PictureClip1.GraphicCell(0)
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = PictureClip1.GraphicCell(1)
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then End
    Dim i As Integer
    Image1.Picture = PictureClip1.GraphicCell(0)
    If Me.Top = -Image1.Top Then
        Do While Me.Top < 0
            Me.Top = Me.Top + 1
            i = 0
            Do While i < 7000
                i = i + 1
            Loop
            Me.Refresh
        Loop
    Else
        Do While Me.Top > -Image1.Top
            Me.Top = Me.Top - 1
            i = 0
            Do While i < 7000
                i = i + 1
            Loop
            Me.Refresh
        Loop
    End If
End Sub
