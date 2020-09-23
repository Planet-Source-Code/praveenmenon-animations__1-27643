VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form13"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   LinkTopic       =   "Form13"
   ScaleHeight     =   3195
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2025
      ScaleWidth      =   6825
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         Height          =   345
         Left            =   5790
         TabIndex        =   1
         Top             =   1620
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000004&
         BorderWidth     =   2
         Height          =   1965
         Index           =   0
         Left            =   30
         Top             =   30
         Width           =   6765
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   2025
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   6825
      End
   End
   Begin VB.Timer timScroll 
      Interval        =   1
      Left            =   6480
      Top             =   2280
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SelectClipPath Lib "gdi32" (ByVal hdc As Long, ByVal iMode As Long) As Long
Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function PathToRegion Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal c As Long, ByVal OP As Long, ByVal cp As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Const FW_DONTCARE = 0
Private Const FW_THIN = 100
Private Const FW_EXTRALIGHT = 200
Private Const FW_ULTRALIGHT = 200
Private Const FW_LIGHT = 300
Private Const FW_NORMAL = 400
Private Const FW_REGULAR = 400
Private Const FW_MEDIUM = 500
Private Const FW_SEMIBOLD = 600
Private Const FW_DEMIBOLD = 600
Private Const FW_BOLD = 700
Private Const FW_EXTRABOLD = 800
Private Const FW_ULTRABOLD = 800
Private Const FW_HEAVY = 900
Private Const FW_BLACK = 900
Private Const RGN_OR = 2

Private Type Size
    cx As Long
    cy As Long
End Type
Private Declare Function GetTextExtentPoint Lib "gdi32" Alias "GetTextExtentPointA" (ByVal hdc As Long, ByVal lpszString As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function StrokePath Lib "gdi32" (ByVal hdc As Long) As Long

Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Private Sub dragMe(frm As Form)
    ReleaseCapture
    Call SendMessage(frm.hWnd, &HA1, 2, 0&)
    
End Sub

' Shape the about window.
Private Sub ShapeForm()
Const Text1 = "Animations"
Const Text2 = ""
Const TEXT_HGT1 = 250 / 1.5
Const TEXT_WID1 = 100 / 1.5
Const TEXT_HGT2 = 250 / 2
Const TEXT_WID2 = 100 / 2
Const FONT_NAME1 = "Impact"
Const FONT_NAME2 = "Times New Roman"
Const VGAP1 = -20
Const VGAP2 = -40
Const DRAW_WIDTH = 7

Dim font1 As Long
Dim font2 As Long
Dim origfont As Long
Dim hRgn1 As Long
Dim hRgn2 As Long
Dim hRgn3 As Long
Dim X1 As Long
Dim X2 As Long
Dim Y1 As Long
Dim Y2 As Long
Dim sz As Size
Dim tm As TEXTMETRIC
Dim wID As Single
Dim hgt As Single
Dim pix_wid As Single
Dim pix_hgt As Single
Dim text1_wid As Single
Dim text1_hgt As Single
Dim text1_int As Single
Dim text2_wid As Single
Dim text2_hgt As Single
Dim text2_int As Single

    ' Prepare the form.
    AutoRedraw = True
    ScaleMode = vbPixels
    BorderStyle = vbBSNone
    BackColor = vbBlue
    ForeColor = vbBlack
    Caption = ""
    DrawWidth = DRAW_WIDTH
    ' ControlBox = False    ' Set at design time.
    ' MinButton = False     ' Set at design time.
    ' MaxButton = False     ' Set at design time.
    ' ShowInTaskbar = False ' Set at design time.

    ' Get the size of the text.
    font1 = CustomFont(TEXT_HGT1, TEXT_WID1, 0, 0, _
        FW_BOLD, False, False, False, _
        FONT_NAME1)
    origfont = SelectObject(hdc, font1)
    GetTextExtentPoint hdc, Text1, Len(Text1), sz
    text1_wid = sz.cx
    GetTextMetrics hdc, tm
    text1_int = tm.tmInternalLeading
    text1_hgt = tm.tmAscent - text1_int

    font2 = CustomFont(TEXT_HGT2, TEXT_WID2, 0, 0, _
        FW_BOLD, False, False, False, _
        FONT_NAME2)
    SelectObject hdc, font1
    GetTextExtentPoint hdc, Text2, Len(Text2), sz
    text2_wid = sz.cx
    GetTextMetrics hdc, tm
    text2_int = tm.tmInternalLeading
    text2_hgt = tm.tmAscent - text2_int

    ' Make the form big enough.
    wID = picLogin.Height
    If wID < text1_wid Then wID = text1_wid
    If wID < text2_wid Then wID = text2_wid
    hgt = picLogin.Height + text1_hgt + text2_hgt + VGAP1 + VGAP2
    pix_wid = ScaleX(wID, vbPixels, vbTwips)
    pix_hgt = ScaleY(hgt, vbPixels, vbTwips)
    Move (Screen.Width - pix_wid) / 2, _
         (Screen.Height - pix_hgt) / 2, _
         pix_wid, pix_hgt

    ' Make the regions.
    SelectObject hdc, font1
    BeginPath hdc
    CurrentX = (wID - text1_wid) / 2
    CurrentY = -text1_int
    Print Text1
    EndPath hdc
    hRgn1 = PathToRegion(hdc)

    SelectObject hdc, font2
    BeginPath hdc
    CurrentX = (wID - text2_wid) / 2
    CurrentY = text1_hgt + VGAP1 + VGAP2 + picLogin.Height - text2_int
    Print Text2
    EndPath hdc
    hRgn2 = PathToRegion(hdc)

    picLogin.Move (wID - picLogin.Width) / 2, text1_hgt + VGAP1
    X1 = picLogin.Left
    X2 = X1 + picLogin.Width
    Y1 = picLogin.Top
    Y2 = Y1 + picLogin.Height
    hRgn3 = CreateRectRgn(X1, Y1, X2, Y2)

    ' Combine the regions.
    CombineRgn hRgn1, hRgn1, hRgn2, RGN_OR
    CombineRgn hRgn1, hRgn1, hRgn3, RGN_OR

    ' Constrain the form to the combined region.
    SetWindowRgn hWnd, hRgn1, False

    ' Draw with a hollow font.
    SelectObject hdc, font1
    BeginPath hdc
    CurrentX = (wID - text1_wid) / 2
    CurrentY = -text1_int
    Print Text1
    EndPath hdc
    StrokePath hdc

    SelectObject hdc, font2
    BeginPath hdc
    CurrentX = (wID - text2_wid) / 2
    CurrentY = text1_hgt + VGAP1 + VGAP2 + picLogin.Height - text2_int
    Print Text2
    EndPath hdc
    StrokePath hdc

    ' Restore the original font.
    SelectObject hdc, origfont

    ' Free font resources (important!)
    DeleteObject font1
    DeleteObject font2
End Sub

' Make a customized font and return its handle.
Private Function CustomFont(ByVal hgt As Long, ByVal wID As Long, ByVal escapement As Long, ByVal orientation As Long, ByVal wgt As Long, ByVal is_italic As Long, ByVal is_underscored As Long, ByVal is_striken_out As Long, ByVal face As String) As Long
Const CLIP_LH_ANGLES = 16   ' Needed for tilted fonts.

    CustomFont = CreateFont( _
        hgt, wID, escapement, orientation, wgt, _
        is_italic, is_underscored, is_striken_out, _
        0, 0, CLIP_LH_ANGLES, 0, 0, face)
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
Form1.Show
Set Form2 = Nothing
End Sub

Private Sub Form_Load()
    ' Shape the form.
    ShapeForm

Dim NextLine As String
Dim GotInfo As String

'display the background colour

NextLine = Chr(13) & Chr(10)
'set the display text
GotInfo = "Project Animations"
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & "Enjoy the animations"
GotInfo = GotInfo & NextLine & "All the source code included are freely usable."
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & "Credits"
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & "Eric O'Sullivan for Scrolling credits"
GotInfo = GotInfo & NextLine & "£ºWXJ_Lake for the form shaping code"
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & "Contact"
GotInfo = GotInfo & NextLine & "praveenc_1999@yahoo.com"
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & "Programmer"
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & "Praveen Menon"
GotInfo = GotInfo & NextLine & "Kerala, India"
GotInfo = GotInfo & NextLine & "*"
Call EnterText(GotInfo)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
dragMe Me
End Sub

Private Sub picLogin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
dragMe Me
End Sub

Private Sub timScroll_Timer()
'scroll the credits up wards.

'in nanoseconds
Const TimePerPixel = 20

Dim Speed As Integer
Static BackArea As Rect
Static Tick As Long

If Tick = 0 Then
    'set the co-ordinates of the background
    picLogin.Cls
    BackArea.Top = 0
    BackArea.Left = 0
    BackArea.Right = (picLogin.ScaleWidth / Screen.TwipsPerPixelX)
    BackArea.Bottom = (picLogin.ScaleHeight / Screen.TwipsPerPixelY)
    
    'set the background of the text
    Call LoadOldBack(picLogin, BackArea)
    Tick = GetTickCount
End If

'if X nanoseconds have elapsed, move text up one pixel
If (Tick + TimePerPixel) < GetTickCount Then
    'move one pixel at a time
    Speed = Screen.TwipsPerPixelY
    
    'move the text up one pixel (Speed)
    Call MoveText(picLogin, BackArea, Speed, vbCentreAlign)
    
    'wait until you can move the text up one pixel
    Tick = GetTickCount
End If
End Sub
