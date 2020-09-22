VERSION 5.00
Begin VB.UserControl u_MenuXP 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D7DBDF&
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   Picture         =   "u_MenuXP.ctx":0000
   ScaleHeight     =   345
   ScaleWidth      =   4800
   ToolboxBitmap   =   "u_MenuXP.ctx":0502
   Begin VB.PictureBox PicShadow 
      Appearance      =   0  'Flat
      BackColor       =   &H00D7DBDF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   4695
      Picture         =   "u_MenuXP.ctx":0814
      ScaleHeight     =   405
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   30
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00D7DBDF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   0
      Left            =   150
      ScaleHeight     =   405
      ScaleWidth      =   765
      TabIndex        =   2
      Tag             =   "1"
      Top             =   30
      Width           =   765
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Menu"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   30
         Width           =   405
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00D7DBDF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   1
      Left            =   915
      ScaleHeight     =   405
      ScaleWidth      =   735
      TabIndex        =   0
      Tag             =   "6"
      Top             =   30
      Width           =   735
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Test"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   30
         Width           =   345
      End
   End
End
Attribute VB_Name = "u_MenuXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Picture1_MouseDown(Index, Button, Shift, x, y)
End Sub

Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim pt As POINTAPI, pts As POINTAPI

Picture1(Index).BackColor = &HDBE3E7: PicShadow.Left = Picture1(Index).Left + Picture1(Index).Width: PicShadow.Visible = True: DoEvents

pt.x = UserControl.ScaleX(Picture1(Index).Left, vbTwips, vbPixels)
pt.y = UserControl.ScaleY(Picture1(Index).Top + Picture1(Index).Height - 130, vbTwips, vbPixels)
ClientToScreen UserControl.hWnd, pt

bMenuWidth = (Picture1(Index).Width / 15) - 1: bTopMenu = True
Picture1(Index).Line (1, Picture1(Index).Height)-(1, 1), RGB(102, 102, 102)
Picture1(Index).Line (1, 1)-(Picture1(Index).Width - 15, 1), RGB(102, 102, 102)
Picture1(Index).Line (Picture1(Index).Width - 15, 1)-(Picture1(Index).Width - 15, Picture1(Index).Height), RGB(102, 102, 102)

'pts.x = UserControl.ScaleX(PicShadow.Left, vbTwips, vbPixels)
'pts.y = UserControl.ScaleY(PicShadow.Top, vbTwips, vbPixels)
'ClientToScreen UserControl.hWnd, pts

'Call CShadowDraw(PicShadow.hWnd, PicShadow.hDC, pts.x, pts.y)

Ret = TrackPopMenu(Picture1(Index).Tag, pt.x, pt.y)

Picture1(Index).BackColor = &HD7DBDF: PicShadow.Visible = False: bTopMenu = False: DoEvents

End Sub

Private Sub UserControl_Initialize()
On Error GoTo ide
    
    bIde = False: Debug.Print 1 / 0
    If bIde <> True Then Call lProcWnd(UserControl.hWnd, True)
    
    Exit Sub
    
ide:

    bIde = True
    Resume Next

End Sub

Private Sub UserControl_Terminate()

    Call lProcWnd(UserControl.hWnd, False)

End Sub

Public Function TrackPopMenu(cMenu As Long, x As Long, y As Long) As Boolean

If bIde = True Then Exit Function
    
Call CInitMenu
Call CSetupMenu(UserControl.hWnd)

'TrackPopupMenuEx Caps(1, 5), TPM_LEFTALIGN Or TPM_TOPALIGN Or TPM_LEFTBUTTON, x, y, UserControl.hWnd, ByVal 0&
TrackPopupMenuEx Caps(cMenu, 5), TPM_LEFTALIGN Or TPM_TOPALIGN Or TPM_LEFTBUTTON, x, y, UserControl.hWnd, ByVal 0&

End Function
