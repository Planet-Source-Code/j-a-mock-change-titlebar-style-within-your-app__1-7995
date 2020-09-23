VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Custom Titlebar"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4560
   Icon            =   "Titlebar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   109
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      Caption         =   "Enable/Disable"
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2775
      Begin VB.CommandButton Sysbut 
         Caption         =   "Close"
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Sysbut 
         Caption         =   "Restore"
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Sysbut 
         Caption         =   "Minimize"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Menu visible"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Value           =   1  'Aktiviert
      Width           =   1215
   End
   Begin VB.PictureBox TitleBar 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000002&
      BorderStyle     =   0  'Kein
      FillStyle       =   0  'Ausgefüllt
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   0
      ScaleHeight     =   18
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   297
      TabIndex        =   0
      Top             =   1200
      Width           =   4455
      Begin VB.CommandButton sMinimize 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   6.75
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3720
         Style           =   1  'Grafisch
         TabIndex        =   3
         Top             =   30
         Width           =   240
      End
      Begin VB.CommandButton sRestore 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   6.75
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3960
         Style           =   1  'Grafisch
         TabIndex        =   2
         Top             =   30
         Width           =   240
      End
      Begin VB.CommandButton sClose 
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   6.75
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4200
         Style           =   1  'Grafisch
         TabIndex        =   1
         Top             =   30
         Width           =   240
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const SWW_HPARENT = (-8)
Private Const WM_MOVE = &H3
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const WM_SIZE = &H5
Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Enum ShowCommands
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_NORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_MAXIMIZE = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
    SW_MAX = 10
End Enum

Private Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, _
    ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
    ByVal Y2 As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, _
    ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" _
    (ByVal hObject As Long) As Long
Private Declare Function GetClipRgn Lib "gdi32" (ByVal hdc As Long, _
    ByVal hRgn As Long) As Long
Private Declare Function OffsetClipRgn Lib "gdi32" (ByVal hdc As Long, _
    ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, _
    lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_VCENTER = &H4
Private Const DT_TOP = &H0
Private Const DT_SINGLELINE = &H20
Private Const DT_RIGHT = &H2
Private Const DT_WORDBREAK = &H10
Private Const DT_CALCRECT = &H400
Private Const DT_WORD_ELLIPSIS = &H40000

Private Type DRAWITEMSTRUCT
   CtlType As Long
   CtlID As Long
   itemID As Long
   itemAction As Long
   itemState As Long
   hwndItem As Long
   hdc As Long
   rcItem As RECT
   itemData As Long
End Type

Public Enum SysMet
    SM_CXSCREEN = 0
    SM_CYSCREEN = 1
    SM_CXVSCROLL = 2
    SM_CYHSCROLL = 3
    SM_CYCAPTION = 4
    SM_CXBORDER = 5
    SM_CYBORDER = 6
    SM_CXDLGFRAME = 7
    SM_CYDLGFRAME = 8
    SM_CYVTHUMB = 9
    SM_CXHTHUMB = 10
    SM_CXICON = 11
    SM_CYICON = 12
    SM_CXCURSOR = 13
    SM_CYCURSOR = 14
    SM_CYMENU = 15
    SM_CXFULLSCREEN = 16
    SM_CYFULLSCREEN = 17
    SM_CYKANJIWINDOW = 18
    SM_MOUSEPRESENT = 19
    SM_CYVSCROLL = 20
    SM_CXHSCROLL = 21
    SM_DEBUG = 22
    SM_SWAPBUTTON = 23
    SM_RESERVED1 = 24
    SM_RESERVED2 = 25
    SM_RESERVED3 = 26
    SM_RESERVED4 = 27
    SM_CXMIN = 28
    SM_CYMIN = 29
    SM_CXSIZE = 30
    SM_CYSIZE = 31
    SM_CXFRAME = 32
    SM_CYFRAME = 33
    SM_CXMINTRACK = 34
    SM_CYMINTRACK = 35
    SM_CXDOUBLECLK = 36
    SM_CYDOUBLECLK = 37
    SM_CXICONSPACING = 38
    SM_CYICONSPACING = 39
    SM_MENUDROPALIGNMENT = 40
    SM_PENWINDOWS = 41
    SM_DBCSENABLED = 42
    SM_CMOUSEBUTTONS = 43
    SM_CMETRICS = 44
End Enum

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const WM_GETSYSMENU = &H313
Private Const WM_NCPAINT = &H85
Private Const WM_DRAWITEM = &H2B
Private Const WM_ACTIVATE = &H6
Dim IsInFocus As Boolean
Dim xad As Long

Implements ISubclass
Private m_emr As EMsgResponse

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Const COLOR_ACTIVECAPTION = 2
Private Const COLOR_INACTIVECAPTION = 3
Private Const COLOR_INACTIVECAPTIONTEXT = 19
Dim colorActive As Long
Dim colorInActive As Long

Private Function MAKELONG(wLow As Long, wHigh As Long) As Long
    MAKELONG = LOWORD(wLow) Or (&H10000 * LOWORD(wHigh))
End Function

Private Function HiWord(ByVal l As Long) As Long
    l = l \ &H10000
    HiWord = Val("&H" & Hex$(l))
End Function

Private Function LOWORD(dwValue As Long) As Long
    CopyMemory LOWORD, dwValue, 2
End Function

Private Sub Check1_Click()
Dim cc As Object
If Check1.Value = 1 Then
  For Each cc In Me
  If TypeOf cc Is Menu Then
    cc.Visible = True
  End If
  Next
Else
  For Each cc In Me
  If TypeOf cc Is Menu Then
    cc.Visible = False
  End If
  Next
End If
End Sub

Private Sub Form_Load()
colorActive = GetSysColor(COLOR_INACTIVECAPTIONTEXT)
colorInActive = GetSysColor(COLOR_INACTIVECAPTION)

TitleBar.ForeColor = colorActive

SetParent TitleBar.hwnd, 0
SetWindowLong TitleBar.hwnd, SWW_HPARENT, Me.hwnd

AttachMessage Me, hwnd, WM_MOVE
AttachMessage Me, hwnd, WM_SIZE
AttachMessage Me, hwnd, WM_NCPAINT
AttachMessage Me, hwnd, WM_ACTIVATE
AttachMessage Me, TitleBar.hwnd, WM_MOVE
AttachMessage Me, TitleBar.hwnd, WM_DRAWITEM

'U can throw the next 2 lines out
'But u also must throw out the 2 lines in the WindowProc Event
'Line (0, 0)-(ScaleWidth, 0), vb3DShadow
'Line (0, 1)-(ScaleWidth, 1), vb3DHighlight
End Sub

Private Sub Form_Unload(Cancel As Integer)
SetWindowLong TitleBar.hwnd, SWW_HPARENT, 0
DetachMessage Me, hwnd, WM_MOVE
DetachMessage Me, hwnd, WM_SIZE
DetachMessage Me, hwnd, WM_NCPAINT
DetachMessage Me, hwnd, WM_ACTIVATE
DetachMessage Me, TitleBar.hwnd, WM_MOVE
DetachMessage Me, TitleBar.hwnd, WM_DRAWITEM
End Sub

Private Property Let ISubClass_MsgResponse(ByVal RHS As EMsgResponse)
m_emr = RHS
End Property

Private Property Get ISubClass_MsgResponse() As EMsgResponse
ISubClass_MsgResponse = m_emr
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim xyFrame As Long
Dim tmpWrct As RECT
Dim tbHeight As Long
GetWindowRect Me.hwnd, tmpWrct
xyFrame = GetSystemMetrics(SysMet.SM_CXFRAME)
tbHeight = GetSystemMetrics(SysMet.SM_CYCAPTION)

If iMsg = WM_MOVE And hwnd = Me.hwnd Then
SetWindowPos TitleBar.hwnd, _
    0, LOWORD(lParam), _
    Top / 15 + xyFrame, _
    tmpWrct.Right - tmpWrct.Left - xyFrame * 2, tbHeight - 1, 0
    
ElseIf iMsg = WM_SIZE And hwnd = Me.hwnd Then
    'U can throw the next 2 lines out
    'Throw these out
    'Line (0, 0)-(ScaleWidth, 0), vb3DShadow
    'Line (0, 1)-(ScaleWidth, 1), vb3DHighlight
    TitleBar.Width = Width - xyFrame * 30
    TitleBar_Paint
ElseIf iMsg = WM_NCPAINT Then
    Dim tmpDC As Long
    Dim hRgn As Long

    tmpDC = GetWindowDC(Me.hwnd)
    
    With tmpWrct
    hRgn = CreateRectRgn(.Left, .Top, .Right, .Bottom)
    End With
    
    ExcludeClipRect tmpDC, xyFrame, xyFrame, tmpWrct.Right - tmpWrct.Left - 4, xyFrame + tbHeight - 1
    
    OffsetClipRgn tmpDC, tmpWrct.Left, tmpWrct.Top
    
    GetClipRgn tmpDC, hRgn
    
    'tmpDC = GetWindowDC(Me.hWnd)
    'ISubclass_WindowProc=
    ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, hRgn, lParam)
    DeleteObject hRgn
ElseIf hwnd = TitleBar.hwnd And iMsg = WM_DRAWITEM Then

    Dim tDis As DRAWITEMSTRUCT
    Dim DrawSty As Long
    CopyMemory tDis, ByVal lParam, Len(tDis)

If tDis.hwndItem = sClose.hwnd Then
    DrawSty = &H0
ElseIf tDis.hwndItem = sRestore.hwnd Then
    If WindowState = vbNormal Then
        DrawSty = &H2
    Else
        DrawSty = &H3
    End If
ElseIf tDis.hwndItem = sMinimize.hwnd Then
    DrawSty = &H1
End If
    
'Debug.Print tDis.itemState
If tDis.itemState = 1 Then
    DrawSty = DrawSty Or &H200
ElseIf tDis.itemState = 4 Then
    DrawSty = DrawSty Or &H100
End If
    DrawFrameControl tDis.hdc, tDis.rcItem, 1, DrawSty

ElseIf iMsg = WM_ACTIVATE And hwnd = Me.hwnd Then
If wParam = 1 Or wParam = 2 Then
    IsInFocus = True
Else
    IsInFocus = False
End If
TitleBar_Paint
'Debug.Print wParam & " " & lParam
'ISubClass_MsgResponse = emrPostProcess
End If

sClose.Width = tbHeight - 3
sRestore.Width = tbHeight - 3
sMinimize.Width = tbHeight - 3

sClose.Left = ScaleWidth - sClose.Width - xyFrame / 2
sRestore.Left = ScaleWidth - (sRestore.Width * 2) - xyFrame / 2 - 2
sMinimize.Left = ScaleWidth - (sMinimize.Width * 3) - xyFrame / 2 - 2

sClose.Height = tbHeight - 5
sRestore.Height = tbHeight - 5
sMinimize.Height = tbHeight - 5
End Function

Private Sub sClose_Click()
Unload Me
End Sub

Private Sub sMinimize_Click()
ShowWindow hwnd, ShowCommands.SW_MINIMIZE
End Sub

Private Sub sRestore_Click()
If WindowState = vbNormal Then
    ShowWindow hwnd, ShowCommands.SW_MAXIMIZE
Else
    ShowWindow hwnd, ShowCommands.SW_NORMAL
End If
sRestore.Refresh
End Sub

Private Sub Sysbut_Click(Index As Integer)
If Index = 0 Then
    If sClose.Enabled Then
        sClose.Enabled = False
    Else
        sClose.Enabled = True
    End If
ElseIf Index = 1 Then
    If sRestore.Enabled Then
        sRestore.Enabled = False
    Else
        sRestore.Enabled = True
    End If
ElseIf Index = 2 Then
    If sMinimize.Enabled Then
        sMinimize.Enabled = False
    Else
        sMinimize.Enabled = True
    End If
End If
End Sub

Sub Gradiate(vPic As PictureBox)
Dim intLoop As Integer
For intLoop = 0 To 255 'Change the colors to get more Effects
  vPic.Line (intLoop, 0)-(vPic.Width, intLoop + 1), RGB(64, 64, 0 + intLoop), BF
Next intLoop
End Sub

Sub Gradiate2(vPic As PictureBox)
Dim intLoop As Integer
For intLoop = 0 To 255 'Change the colors to get more Effects
  vPic.Line (intLoop, 0)-(Screen.Width, intLoop + 1), RGB(64, 64, 255 - intLoop), BF
Next intLoop
End Sub

Private Sub TitleBar_DblClick()
Dim xyC As POINTAPI
Dim xyFrame As Long
xyFrame = GetSystemMetrics(SysMet.SM_CXFRAME)
GetCursorPos xyC

If xyC.x - (Left / 15) - xyFrame > TitleBar.ScaleHeight And _
    xyC.x - (Left / 15) - xyFrame < sMinimize.Left - 2 Then
Call sRestore_Click
ElseIf xyC.x - (Left / 15) - xyFrame < TitleBar.ScaleHeight Then
    Unload Me
End If
End Sub

Private Sub TitleBar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.SetFocus

Dim xyFrame As Long
Dim tbHeight As Long
xyFrame = GetSystemMetrics(SysMet.SM_CXBORDER)
tbHeight = GetSystemMetrics(SysMet.SM_CYCAPTION)

If x < TitleBar.ScaleHeight And Button = 1 Then
    Dim tmpC As POINTAPI
    GetCursorPos tmpC
    SendMessage hwnd, WM_GETSYSMENU, 0, ByVal MAKELONG(Left / 15 + xyFrame * 2 + 2, Top / 15 + xyFrame + tbHeight + 2)
End If
End Sub

Private Sub TitleBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    Call ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Sub TitleBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    Call ReleaseCapture
ElseIf Button = 2 Then
    Dim tmpC As POINTAPI
    GetCursorPos tmpC
    SendMessage hwnd, WM_GETSYSMENU, 0, ByVal MAKELONG(tmpC.x, tmpC.y)
End If
End Sub

Private Sub TitleBar_Paint()
Dim TheIcon As Long
Dim xx, yy As Long
TheIcon = Me.Icon

With TitleBar
.Cls

If IsInFocus Then
  TitleBar.ForeColor = colorActive
  Gradiate TitleBar
Else
  TitleBar.ForeColor = colorInActive
  Gradiate2 TitleBar
End If

DrawIconEx .hdc, 1, 1, TheIcon, .ScaleHeight - 2, .ScaleHeight - 2, ByVal 0&, ByVal 0&, &H3

Dim sText As String
Dim tmprect As RECT
Dim rcItem As RECT
sText = Me.Caption

tmprect.Left = 0
tmprect.Right = .ScaleWidth - .ScaleHeight - (.ScaleWidth - sMinimize.Left) - 10
tmprect.Bottom = .ScaleHeight

    DrawText .hdc, sText, -1, tmprect, DT_LEFT Or _
    DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_CALCRECT

tmprect.Top = (.ScaleHeight / 2) - (tmprect.Bottom / 2)
tmprect.Bottom = tmprect.Bottom + tmprect.Top
tmprect.Left = .ScaleHeight + 5
tmprect.Right = tmprect.Right + tmprect.Left

    DrawText .hdc, sText, -1, tmprect, DT_LEFT Or _
    DT_SINGLELINE Or DT_WORD_ELLIPSIS
.Refresh
End With
End Sub
