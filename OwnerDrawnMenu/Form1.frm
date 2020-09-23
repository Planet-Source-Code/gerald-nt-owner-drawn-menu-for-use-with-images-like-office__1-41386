VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   2355
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   2355
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Show MsgBox"
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.PictureBox picAlternative 
      BorderStyle     =   0  'Kein
      Height          =   255
      Left            =   2280
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   1920
      Width           =   255
   End
   Begin VB.PictureBox picSeparator 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   15
      Left            =   1440
      Picture         =   "Form1.frx":00D5
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.PictureBox picUnchecked 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1800
      Picture         =   "Form1.frx":010C
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   4
      Top             =   1920
      Width           =   300
   End
   Begin VB.PictureBox picMainBackground 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   0
      Picture         =   "Form1.frx":017F
      ScaleHeight     =   225
      ScaleWidth      =   1200
      TabIndex        =   3
      Top             =   1440
      Width           =   1200
   End
   Begin VB.PictureBox picCheck 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   300
      Left            =   1440
      Picture         =   "Form1.frx":085E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   2
      Top             =   1920
      Width           =   300
   End
   Begin VB.PictureBox picSelect 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'Kein
      Height          =   540
      Left            =   1320
      Picture         =   "Form1.frx":0921
      ScaleHeight     =   540
      ScaleWidth      =   1695
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.PictureBox picBackground 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   390
      Left            =   120
      Picture         =   "Form1.frx":0BF4
      ScaleHeight     =   390
      ScaleWidth      =   690
      TabIndex        =   0
      Top             =   1800
      Width           =   690
   End
   Begin VB.Menu mnu1 
      Caption         =   "Menu1"
      Begin VB.Menu mnu3 
         Caption         =   "&Menu3"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnu4 
         Caption         =   "Menu4"
      End
      Begin VB.Menu mnu5 
         Caption         =   "Menu5"
      End
      Begin VB.Menu mnu6 
         Caption         =   "Menu6"
      End
      Begin VB.Menu mnu7 
         Caption         =   "Menu7"
      End
   End
   Begin VB.Menu mnu20 
      Caption         =   "Menu20"
      Begin VB.Menu mnu21 
         Caption         =   "Menu21"
         Begin VB.Menu mnu22Sub 
            Caption         =   "Menu22 sub"
         End
      End
      Begin VB.Menu mnu23 
         Caption         =   "Mnu23"
      End
      Begin VB.Menu mnu24 
         Caption         =   "Mnu24"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim menu As CMenu

Private Sub Command1_Click()
    MsgBox "MsgBox()"
End Sub

Private Sub Form_Load()

    Set menu = New CMenu
    
    'subclass window
    modSubclassing.SubClass Me.hWnd, ObjPtr(Me), AddressOf RedirectWndProc

    'init menu - object
    menu.init Me.hWnd, picMainBackground, picBackground, picSelect, picCheck, picUnchecked, picSeparator

    'overwrite standard-images
    menu.setItemProperties 3, picUnchecked:=picAlternative

    'or set it nothing to use windows-standard
    menu.setItemProperties 4, , Nothing
    menu.setItemProperties 6, picSelect, Nothing, , Nothing
    menu.setItemProperties 8, picUnchecked:=Nothing

End Sub

Friend Function WndProc(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    menu.processMessage hWnd, lMsg, wParam, lParam
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnu1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    modSubclassing.UnSubClass
End Sub

Private Sub mnu22Sub_Click()
    mnu22Sub.Checked = Not mnu22Sub.Checked
End Sub

Private Sub mnu23_Click()
    mnu23.Checked = Not mnu23.Checked
End Sub

Private Sub mnu24_Click()
    mnu24.Checked = Not mnu24.Checked
End Sub

Private Sub mnu3_Click()
    mnu3.Checked = Not mnu3.Checked
End Sub

Private Sub mnu4_Click()
    mnu4.Checked = Not mnu4.Checked
End Sub

Private Sub mnu6_Click()
    mnu6.Checked = Not mnu6.Checked
End Sub

Private Sub mnu7_Click()
    mnu7.Checked = Not mnu7.Checked
End Sub

