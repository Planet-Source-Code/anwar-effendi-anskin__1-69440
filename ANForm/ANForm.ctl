VERSION 5.00
Begin VB.UserControl ANForm 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BackStyle       =   0  'Transparent
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   345
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00800000&
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   345
   ScaleMode       =   0  'User
   ScaleWidth      =   345
   Tag             =   "ANForm"
   ToolboxBitmap   =   "ANForm.ctx":0000
   Begin VB.PictureBox picBottomBorder 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1800
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   3
      Top             =   1320
      Width           =   225
   End
   Begin VB.PictureBox picRightBorder 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2040
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   1080
      Width           =   225
   End
   Begin VB.PictureBox picLeftBorder 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2040
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   1320
      Width           =   225
   End
   Begin VB.PictureBox picTitleBar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1800
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   1080
      Width           =   225
   End
   Begin VB.Image imgLogo 
      Height          =   350
      Left            =   0
      Picture         =   "ANForm.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   350
   End
End
Attribute VB_Name = "ANForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'************************************************************************************
'*******************************  API  Section  ***********************************
'************************************************************************************

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Boolean, ByVal fdwUnderline As Boolean, ByVal fdwStrikeOut As Boolean, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub ReleaseCapture Lib "user32" ()

Private Const SWW_HPARENT = (-8)
Private Const TRANSPARENT = 1
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const DT_END_ELLIPSIS = &H8000&

Private Const LOGPIXELSY = 90
Private Const FW_BOLD = 700
Private Const DEFAULT_CHARSET = 1
Private Const OUT_DEFAULT_PRECIS = 0
Private Const PROOF_QUALITY = 2
Private Const DEFAULT_PITCH = 0
Private Const CLIP_DEFAULT_PRECIS = 0

Private Const GWL_WNDPROC = (-4)
Private Const SM_CXFRAME = 32
Private Const SM_CYFRAME = 33
Private Const SM_CXSIZE = 30
Private Const SM_CYCAPTION = 4
Private Const SM_CYMENU = 15
Private Const SM_CXBORDER = 5
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const WM_GETSYSMENU = &H313

Private Const SRCCOPY = &HCC0020

Private Type MyColor
    R As Integer
    G As Integer
    B As Integer
End Type

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

Public Enum ColorSkinStyle
    BlackWhiteStyle = 0
    RedStyle = 1
    OrangeStyle = 2
    GreenStyle = 3
    CyanStyle = 4
    BlueStyle = 5
    MagentaStyle = 6
End Enum

Private frmParent As Form
Attribute frmParent.VB_VarHelpID = -1
Private intCursorX As Integer, intCursorY As Integer
Private m_StyleColor As ColorSkinStyle
Private m_FormActive As Boolean
Private WithEvents frmUnload As Form
Attribute frmUnload.VB_VarHelpID = -1

'************************************************************************************
'***************************  Sub and Function Events Section ***********************
'************************************************************************************

Private Sub BeginLoadSkin()
Dim borderWidth As Long
Dim TitleBarHeight As Long
Dim TitleBarWidth As Long
Dim MyFont As StdFont
Dim X As Long, Y As Long
Dim myMenu As Object
Set frmParent = UserControl.ParentControls.Item(0)

    SetWindowLong picTitleBar.hwnd, SWW_HPARENT, frmParent.hwnd
    SetParent picTitleBar.hwnd, 0
    SetWindowLong picLeftBorder.hwnd, SWW_HPARENT, frmParent.hwnd
    SetParent picLeftBorder.hwnd, 0
    SetWindowLong picRightBorder.hwnd, SWW_HPARENT, frmParent.hwnd
    SetParent picRightBorder.hwnd, 0
    SetWindowLong picBottomBorder.hwnd, SWW_HPARENT, frmParent.hwnd
    SetParent picBottomBorder.hwnd, 0

'Form Parent Property Set
    frmParent.AutoRedraw = True
    frmParent.BackColor = vbWhite
    frmParent.WindowState = 0
    UserControl.ScaleMode = 3
    
'Replace Original Title Bar with New Title Bar
    borderWidth = GetSystemMetrics(SM_CXFRAME) - 1
    TitleBarWidth = frmParent.ScaleX(frmParent.ScaleWidth, vbTwips, vbPixels) + 3 * borderWidth
    TitleBarHeight = GetSystemMetrics(SM_CYCAPTION)
    X = frmParent.ScaleWidth + ScaleX(borderWidth + 3, vbPixels, vbTwips)
    Y = frmParent.ScaleHeight + ScaleY(TitleBarHeight + 3, vbPixels, vbTwips)
    
    picTitleBar.Width = X
    picTitleBar.Height = 300
    picTitleBar.BackColor = DrawGradientColor(False, 0.2)
    picBottomBorder.Width = X + 5
    picBottomBorder.Height = 50
    
For Each myMenu In frmParent
    If TypeOf myMenu Is Menu Then
        picLeftBorder.Height = Y
        picRightBorder.Height = Y
    Else
        picLeftBorder.Height = frmParent.ScaleHeight
        picRightBorder.Height = frmParent.ScaleHeight
    End If
Next

    picLeftBorder.Width = 50
    picLeftBorder.BackColor = DrawGradientColor(False, 0.2)
    picRightBorder.Width = 50
    picRightBorder.BackColor = DrawGradientColor(False, 0.2)

    picTitleBar.Top = frmParent.Top
    picTitleBar.Left = frmParent.Left
    picLeftBorder.Top = frmParent.Top + picTitleBar.ScaleHeight + borderWidth
    picLeftBorder.Left = frmParent.Left
    picRightBorder.Top = frmParent.Top + picTitleBar.ScaleHeight + borderWidth
    picRightBorder.Left = frmParent.Left + frmParent.Width - 50
    picBottomBorder.Top = frmParent.Top + frmParent.Height - 50
    picBottomBorder.Left = frmParent.Left

    Call DrawPicSkin
End Sub

Private Sub DrawPicSkin()
Dim winDC As Long
Dim borderWidth As Long
Dim TitleBarHeight As Long
Dim myMenu As Object
Dim rct As RECT, skinCaption As String
Dim TitleBarWidth As Long
Dim X As Long, Y As Long, i As Integer, j As Integer, k As Integer, l As Integer

'Determine size of Pictures
    winDC = GetWindowDC(frmParent.hwnd)
    borderWidth = GetSystemMetrics(SM_CXFRAME) - 1
    TitleBarHeight = GetSystemMetrics(SM_CYCAPTION)
    TitleBarWidth = frmParent.ScaleX(frmParent.ScaleWidth, vbTwips, vbPixels) + 3 * borderWidth
    X = frmParent.ScaleWidth + ScaleX(borderWidth + 3, vbPixels, vbTwips)
    Y = frmParent.ScaleHeight + ScaleY(TitleBarHeight + 3, vbPixels, vbTwips)
    
For i = 0 To 50
'Draw Left Menu Border
    picLeftBorder.ForeColor = DrawGradientColor(True, i / 50)
    picLeftBorder.Line (i, i)-(i, Y)
    
'Draw Right Menu Border
    picRightBorder.ForeColor = DrawGradientColor(True, i / 50)
    picRightBorder.Line (i, i)-(i, Y)
    
'Draw Bottom Border
    picBottomBorder.ForeColor = DrawGradientColor(True, i / 50)
    picBottomBorder.Line (0, i)-(X, i)

For Each myMenu In frmParent
    If TypeOf myMenu Is Menu Then
        frmParent.ForeColor = DrawGradientColor(True, i / 50)
        frmParent.Line (0, 0)-(X, 0)
    End If
Next
Next i

'Draw Tittle Bar
For j = 0 To 300
    picTitleBar.ForeColor = DrawGradientColor(True, j / 350)
    picTitleBar.Line (0, j)-(X, j)
    
    picTitleBar.ForeColor = DrawGradientColor(False, j / 200)
    picTitleBar.Line (550 - j, j)-(X - 50 - j, j)
    picTitleBar.ForeColor = DrawGradientColor(True, j / 300)
    picTitleBar.Line (X - 120 - j, j)-(X - 80 - j, j)
    picTitleBar.ForeColor = DrawGradientColor(True, j / 300)
    picTitleBar.Line (X - 170 - j, j)-(X - 150 - j, j)
    picTitleBar.ForeColor = DrawGradientColor(True, j / 300)
    picTitleBar.Line (X - 240 - j, j)-(X - 210 - j, j)
    
    picTitleBar.ForeColor = DrawGradientColor(True, j / 300)
    picTitleBar.Line (580 - j, j)-(600 - j, j)
    picTitleBar.ForeColor = DrawGradientColor(True, j / 300)
    picTitleBar.Line (630 - j, j)-(660 - j, j)
    picTitleBar.ForeColor = DrawGradientColor(True, j / 300)
    picTitleBar.Line (690 - j, j)-(720 - j, j)
Next j

'Draw Tittle Bar Accesories
For k = 0 To 30
    picTitleBar.ForeColor = DrawGradientColor(False, 0.4)
    picTitleBar.Line (X - k, 0)-(X - k, 350)
If k > 30 Then GoTo KLain
    picTitleBar.ForeColor = DrawGradientColor(False, 0.4)
    picTitleBar.Line (k, 0)-(k, 380)
KLain:
Next k

    rct.Top = borderWidth + 1.5
    rct.Left = 50
    rct.Right = TitleBarWidth - frmParent.ScaleX(1300, vbTwips, vbPixels)
    rct.Bottom = TitleBarHeight - borderWidth

    SetBkMode picTitleBar.hdc, TRANSPARENT
    SetTextColor picTitleBar.hdc, DrawGradientColor(False, 0)
    rct.Top = borderWidth + 1.5
    rct.Left = 50
    rct.Right = TitleBarWidth - frmParent.ScaleX(1300, vbTwips, vbPixels)
    rct.Bottom = TitleBarHeight - borderWidth
    skinCaption = frmParent.Caption
    DeleteObject SelectObject(winDC, w_FontWnd(8, "Tahoma"))
    DrawText picTitleBar.hdc, skinCaption, Len(skinCaption), rct, DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_VCENTER
    DrawIconEx picTitleBar.hdc, 6, 2, frmParent.Icon, 16, 16, ByVal 0&, ByVal 0&, &H3
End Sub

Private Function w_FontWnd(nSize As Integer, nFontName As String) As Long
    w_FontWnd = CreateFont(-MulDiv(nSize, GetDeviceCaps(GetDC(0), LOGPIXELSY), 72), 0, 0, 0, FW_BOLD, False, False, False, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, DEFAULT_PITCH, nFontName)
End Function

Private Function ColorRGB(ByVal ColorVal As Long) As MyColor
    ColorRGB.R = ColorVal Mod 256
    ColorRGB.G = ((ColorVal And &HFF00FF00) / 256&)
    ColorRGB.B = (ColorVal And &HFF0000) / (256& * 256&)
End Function

Private Function DrawGradientColor(HeavyToLight As Boolean, Fraction As Currency) As Long
Dim ResultColor As MyColor, Col1 As MyColor, Col2 As MyColor
Dim Color1 As Long, Color2 As Long
If m_FormActive = False Then
        If HeavyToLight = True Then
            Color1 = &HC0C0C0
            Color2 = &HE0E0E0
        Else
            Color1 = &HE0E0E0
            Color2 = &HC0C0C0
        End If
Else
Select Case m_StyleColor
    Case 0
        If HeavyToLight = True Then
            Color1 = &H0&
            Color2 = &HFFFFFF
        Else
            Color1 = &HFFFFFF
            Color2 = &H0&
        End If
    Case 1
        If HeavyToLight = True Then
            Color1 = &HFF&
            Color2 = &HC0C0FF
        Else
            Color1 = &HC0C0FF
            Color2 = &HFF&
        End If
    Case 2
        If HeavyToLight = True Then
            Color1 = &H80FF&
            Color2 = &HC0E0FF
        Else
            Color1 = &HC0E0FF
            Color2 = &H80FF&
        End If
    Case 3
        If HeavyToLight = True Then
            Color1 = &HC000&
            Color2 = &HC0FFC0
        Else
            Color1 = &HC0FFC0
            Color2 = &HC000&
        End If
    Case 4
        If HeavyToLight = True Then
            Color1 = &HC0C000
            Color2 = &HFFFFC0
        Else
            Color1 = &HFFFFC0
            Color2 = &HC0C000
        End If
    Case 5
        If HeavyToLight = True Then
            Color1 = &HC00000
            Color2 = &HFFC0C0
        Else
            Color1 = &HFFC0C0
            Color2 = &HC00000
        End If
    Case 6
        If HeavyToLight = True Then
            Color1 = &HC000C0
            Color2 = &HFFC0FF
        Else
            Color1 = &HFFC0FF
            Color2 = &HC000C0
        End If
End Select
End If
    Col1 = ColorRGB(Color1)
    Col2 = ColorRGB(Color2)
    ResultColor.R = Fraction * (Col1.R - Col2.R) + Col2.R
    ResultColor.G = Fraction * (Col1.G - Col2.G) + Col2.G
    ResultColor.B = Fraction * (Col1.B - Col2.B) + Col2.B
    DrawGradientColor = RGB(ResultColor.R, ResultColor.G, ResultColor.B)
End Function

'************************************************************************************
'****************************  Property Section  ************************************
'************************************************************************************

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_StyleColor = PropBag.ReadProperty("StyleColor", 0)
    m_FormActive = PropBag.ReadProperty("FormActive", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("StyleColor", m_StyleColor, 0)
    Call PropBag.WriteProperty("FormActive", m_FormActive, True)
End Sub

Public Property Get StyleColor() As ColorSkinStyle
    StyleColor = m_StyleColor
End Property

Public Property Let StyleColor(ByVal New_StyleColor As ColorSkinStyle)
    m_StyleColor = New_StyleColor
    PropertyChanged "StyleColor"
End Property

Private Property Get FormActive() As Boolean
    FormActive = m_FormActive
End Property

Private Property Let FormActive(ByVal New_FormActive As Boolean)
    m_FormActive = New_FormActive
    PropertyChanged "FormActive"
End Property

'************************************************************************************
'********************************  Events Section  **********************************
'************************************************************************************

Private Sub UserControl_Resize()
    Call UserControl_Show
End Sub

Private Sub UserControl_Show()
If UserControl.Ambient.UserMode = False Then
    imgLogo.Visible = True
    UserControl.Height = 350
    UserControl.Width = 350
Else
    Call BeginLoadSkin
    imgLogo.Visible = False
    Set frmUnload = frmParent
End If
End Sub

Private Sub frmUnload_Unload(Cancel As Integer)
    SetWindowLong picTitleBar.hwnd, SWW_HPARENT, 0
    SetWindowLong picRightBorder.hwnd, SWW_HPARENT, 0
    SetWindowLong picLeftBorder.hwnd, SWW_HPARENT, 0
    SetWindowLong picBottomBorder.hwnd, SWW_HPARENT, 0
    frmUnload.Hide
Do
    picTitleBar.Top = picTitleBar.Top + 100
    picRightBorder.Left = picRightBorder.Left - 40
    picLeftBorder.Left = picLeftBorder.Left + 40
    picBottomBorder.Top = picBottomBorder.Top - 40
Loop Until picTitleBar.Top > Screen.Height
End Sub

Private Sub frmUnload_Load()
    frmUnload.Top = (Screen.Height - frmUnload.Height) / 4
    frmUnload.Left = (Screen.Width - frmUnload.Width) / 2
    Call BeginLoadSkin
End Sub

Private Sub frmUnload_Deactivate()
    m_FormActive = False
    Call DrawPicSkin
End Sub

Private Sub frmUnload_Activate()
    m_FormActive = True
    Call DrawPicSkin
End Sub

Private Sub picTitlebar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xyFrame As Long
Dim tbHeight As Long
Dim tmpC As POINTAPI

xyFrame = GetSystemMetrics(SM_CXBORDER)
tbHeight = GetSystemMetrics(SM_CYCAPTION)

If X < picTitleBar.ScaleHeight And Button = 1 Then
    GetCursorPos tmpC
    SendMessage frmParent.hwnd, WM_GETSYSMENU, 0, ByVal 0&
End If
End Sub

Private Sub picTitlebar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Call ReleaseCapture
    SendMessage frmParent.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    Call BeginLoadSkin
End If
End Sub
