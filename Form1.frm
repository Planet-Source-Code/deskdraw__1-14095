VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "DeskDecorator"
   ClientHeight    =   1884
   ClientLeft      =   120
   ClientTop       =   576
   ClientWidth     =   3744
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1884
   ScaleWidth      =   3744
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   396
      Left            =   1632
      TabIndex        =   5
      Top             =   1476
      Width           =   972
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   396
      Left            =   336
      TabIndex        =   4
      Top             =   1488
      Width           =   972
   End
   Begin VB.TextBox txtAltPicPath 
      Height          =   312
      Left            =   108
      TabIndex        =   3
      Top             =   780
      Width           =   2436
   End
   Begin VB.TextBox txtPicPath 
      Height          =   312
      Left            =   84
      TabIndex        =   2
      Top             =   240
      Width           =   2436
   End
   Begin VB.PictureBox picAltPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   396
      Left            =   2640
      ScaleHeight     =   348
      ScaleWidth      =   924
      TabIndex        =   1
      Top             =   768
      Width           =   972
   End
   Begin VB.PictureBox picPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   396
      Left            =   2616
      ScaleHeight     =   348
      ScaleWidth      =   924
      TabIndex        =   0
      Top             =   216
      Width           =   972
   End
   Begin VB.Timer Timer1 
      Left            =   672
      Top             =   1176
   End
   Begin VB.Menu MnuPopupMenu 
      Caption         =   "File"
      Begin VB.Menu MnuStart 
         Caption         =   "&Start"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'you can draw images in any DC
'if you have any doubts, mail me tk_pramod@yahoo.com
'you can get more VB source code from my home page www.imprudents.com

Option Explicit
Private Type POINTAPI
    x As Long
    y As Long
End Type
Const LR_LOADFROMFILE = &H10
Const IMAGE_BITMAP = 0
Const IMAGE_ICON = 1
Const IMAGE_CURSOR = 2
Const IMAGE_ENHMETAFILE = 3
Const CF_BITMAP = 2
Const RDW_INVALIDATE = &H1
Const RDW_INTERNALPAINT = &H2
Const BS_HATCHED = 2
Const HS_CROSS = 4
Const SW_SHOWNORMAL = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim return1 As Long, return2 As Long, hBitmap As Long
Dim i As Long, j As Long, MaxI As Long, MaxJ As Long
Dim strPic As String, strAltPic As String
Dim Pos(100) As POINTAPI
Dim cnt As Integer
Private Const APP_NAME = "DeskDecorator"


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
'save settings to registry
SaveSetting APP_NAME, "Configuration", "Pic", strPic
SaveSetting APP_NAME, "Configuration", "AltPic", strAltPic
'loading two images
hBitmap = LoadImage(App.hInstance, strPic, IMAGE_ICON, 32, 32, LR_LOADFROMFILE)
If hBitmap = 0 Then
    MsgBox "There was an error while loading the bitmap"
    End
Else
    return1& = hBitmap
End If
hBitmap = LoadImage(App.hInstance, strAltPic, IMAGE_ICON, 32, 32, LR_LOADFROMFILE)
If hBitmap = 0 Then
    MsgBox "There was an error while loading the bitmap"
    End
Else
    return2& = hBitmap
End If
'gets the screen width & height

MaxI = (Screen.Width / Screen.TwipsPerPixelX) - 64
MaxJ = (Screen.Height / Screen.TwipsPerPixelY) - 64

'stores the x,y axis to the pos array to draw the images
cnt = 0
For i = 1 To MaxI Step 64
    Pos(cnt).x = i
    Pos(cnt).y = 1
    cnt = cnt + 1
Next
For j = 1 To MaxJ Step 64
    Pos(cnt).x = MaxI
    Pos(cnt).y = j
    cnt = cnt + 1
Next
For i = MaxI To 1 Step -64
    Pos(cnt).x = i
    Pos(cnt).y = MaxJ
    cnt = cnt + 1
Next
For j = MaxJ To 1 Step -64
    Pos(cnt).x = 1
    Pos(cnt).y = j
    cnt = cnt + 1
Next



End Sub

Private Sub Form_Load()
'get the old settings from registry
strPic = GetSetting(APP_NAME, "Configuration", "Pic", "")
strAltPic = GetSetting(APP_NAME, "Configuration", "AltPic", "")
If strPic = "" Or strAltPic = "" Then Exit Sub
txtAltPicPath = strAltPic
txtPicPath = strPic
picPic.Picture = LoadPicture(strPic)
picAltPic.Picture = LoadPicture(strAltPic)
Timer1.Interval = 1000
Timer1.Enabled = False
Dim hDC As Long, hBitmap As Long
End Sub

Private Sub Form_Unload(Cancel As Integer)
'before unload, refresh desktop to clear our draws, if any
InvalidateRect 0, 0, 1
End Sub

Private Sub Label2_Click()
ShellExecute Me.hwnd, vbNullString, "http://www.tripodasia.com.sg/Imprudent", vbNullString, "C:\", SW_SHOWNORMAL
End Sub

Private Sub MnuConfigure_Click()
CloseApp
Me.Show
End Sub

Private Sub MnuExit_Click()
Unload Me
End Sub

Private Sub MnuStart_Click()
If MnuStart.Caption = "&Start" Then
    MnuStart.Caption = "&Stop"
    Timer1.Interval = 1000
    Timer1.Enabled = True
Else
    'refresh desktop to clear draws
    InvalidateRect 0, 0, 1
    MnuStart.Caption = "&Start"
    Timer1.Enabled = False
End If
End Sub

Private Sub picAltPic_Click()
strAltPic = ShowOpen(Me.hwnd, "Icon Files  (*.ico)" & Chr(0) & "*.ico", App.Path, "Select Alter Icon")
Me.txtAltPicPath = strAltPic
picAltPic.Picture = LoadPicture(strAltPic)
End Sub

Private Sub picPic_Click()
strPic = ShowOpen(Me.hwnd, "Icon Files  (*.ico)" & Chr(0) & "*.ico", App.Path, "Select Icon")
txtPicPath = strPic
picPic.Picture = LoadPicture(strPic)
End Sub

Sub Timer1_Timer()

Static flg As Boolean
'draws icon on desktop
If flg Then
    For i = 0 To cnt
        DrawIcon GetWindowDC(0), Pos(i).x, Pos(i).y, return1&
    Next
Else
    For i = 0 To cnt
        DrawIcon GetWindowDC(0), Pos(i).x, Pos(i).y, return2&
    Next
End If
flg = Not flg
End Sub
