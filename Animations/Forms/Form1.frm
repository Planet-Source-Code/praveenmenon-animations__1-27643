VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2265
   ClientLeft      =   12990
   ClientTop       =   9465
   ClientWidth     =   2340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   156
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1710
      Top             =   1350
   End
   Begin PicClip.PictureClip PictureClip1 
      Left            =   1440
      Top             =   2400
      _ExtentX        =   73792
      _ExtentY        =   3149
      _Version        =   393216
      Cols            =   6
   End
   Begin VB.Menu Praveen 
      Caption         =   "Praveen"
      Visible         =   0   'False
      Begin VB.Menu mnuHi 
         Caption         =   "Hi"
      End
      Begin VB.Menu mnuHi1 
         Caption         =   "Hi1"
      End
      Begin VB.Menu mnuRing 
         Caption         =   "Ring"
      End
      Begin VB.Menu mnuFlower 
         Caption         =   "Flower"
      End
      Begin VB.Menu mnuTop 
         Caption         =   "RedTop"
      End
      Begin VB.Menu mnuFlame 
         Caption         =   "Flame"
      End
      Begin VB.Menu mnuSmoke 
         Caption         =   "Smoke"
      End
      Begin VB.Menu mnuDove 
         Caption         =   "Dove"
      End
      Begin VB.Menu mnuBat 
         Caption         =   "Bat"
      End
      Begin VB.Menu mnuBFly 
         Caption         =   "ButterFly"
      End
      Begin VB.Menu mnuCheetah 
         Caption         =   "Cheetah"
      End
      Begin VB.Menu mnuFish 
         Caption         =   "Fish"
      End
      Begin VB.Menu mnuDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPraveen 
         Caption         =   "Praveen"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim y As Integer
Dim maxInt As Integer

Private Sub Form_DblClick()

Unload Me

End Sub

Private Sub Form_Load()

PictureClip1.Picture = LoadPicture(App.Path & "\images\hi.bmp")
MsgBox "Right Click on the animation for a context menu", vbInformation
Dim t As Single
Dim rtn As Long
                            'initial adjustments

PictureClip1.Rows = 1
PictureClip1.Cols = 22
maxInt = 22
y = 0

Form1.Picture = PictureClip1.GraphicCell(0)


If Me.Picture <> 0 Then
  Call SetAutoRgn(Me)       'this is the formshaping function
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
                            'This piece of code is
                            'solely for dragging the animation
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
Else
PopupMenu Praveen
End If
End Sub

Private Sub mnuAbout_Click()

'About form displayed
Unload Me
Form2.Show vbModal

End Sub

                            'From here are the changes in the picture property, when
                            'a menu is selected

Private Sub mnuBat_Click()
PictureClip1.Picture = LoadPicture(App.Path & "\images\bat.bmp")
PictureClip1.Rows = 1
PictureClip1.Cols = 10
maxInt = 10
y = 0
End Sub

Private Sub mnuBFly_Click()
PictureClip1.Picture = LoadPicture(App.Path & "\images\bfly.bmp")
PictureClip1.Rows = 1
PictureClip1.Cols = 6
maxInt = 6
y = 0
End Sub

Private Sub mnuCheetah_Click()
PictureClip1.Picture = LoadPicture(App.Path & "\images\cheetah.bmp")
PictureClip1.Rows = 1
PictureClip1.Cols = 8
maxInt = 8
y = 0
End Sub

Private Sub mnuDove_Click()
PictureClip1.Picture = LoadPicture(App.Path & "\images\dove.bmp")
PictureClip1.Rows = 1
PictureClip1.Cols = 18
maxInt = 18
y = 0
End Sub

Private Sub mnuExit_Click()
Unload Form1
Set Form1 = Nothing
End
End Sub

Private Sub mnuFish_Click()
PictureClip1.Picture = LoadPicture(App.Path & "\images\fish.bmp")
PictureClip1.Rows = 1
PictureClip1.Cols = 13
maxInt = 13
y = 0
End Sub

Private Sub mnuFlame_Click()
PictureClip1.Picture = LoadPicture(App.Path & "\images\flame.bmp")
PictureClip1.Rows = 1
PictureClip1.Cols = 8
maxInt = 8
y = 0
End Sub

Private Sub mnuFlower_Click()
PictureClip1.Picture = LoadPicture(App.Path & "\images\flower.bmp")
PictureClip1.Rows = 1
PictureClip1.Cols = 6
maxInt = 6
y = 0
End Sub

Private Sub mnuHi_Click()
PictureClip1.Picture = LoadPicture(App.Path & "\images\hi.bmp")
PictureClip1.Rows = 1
PictureClip1.Cols = 22
maxInt = 22
y = 0
End Sub

Private Sub mnuHi1_Click()
PictureClip1.Picture = LoadPicture(App.Path & "\images\hi1.bmp")
PictureClip1.Rows = 1
PictureClip1.Cols = 22
maxInt = 22
y = 0
End Sub

Private Sub mnuPraveen_Click()
PictureClip1.Picture = LoadPicture(App.Path & "\images\Praveen.bmp")
PictureClip1.Rows = 1
PictureClip1.Cols = 10
maxInt = 10
y = 0
End Sub

Private Sub mnuRing_Click()
PictureClip1.Picture = LoadPicture(App.Path & "\images\ring.bmp")
PictureClip1.Rows = 1
PictureClip1.Cols = 18
maxInt = 18
y = 0
End Sub

Private Sub mnuSmoke_Click()
PictureClip1.Picture = LoadPicture(App.Path & "\images\smoke.bmp")
PictureClip1.Rows = 1
PictureClip1.Cols = 10
maxInt = 10
y = 0
End Sub

Private Sub mnuTop_Click()
PictureClip1.Picture = LoadPicture(App.Path & "\images\top.bmp")
PictureClip1.Rows = 1
PictureClip1.Cols = 9
maxInt = 9
y = 0
End Sub

'The timer code changes the picture property of the form
'by taking the pictureclip's current graphiccell

Private Sub Timer1_Timer()
If y = maxInt - 1 Then
y = 0
Else
y = y + 1
End If
Form1.Refresh
Form1.Picture = PictureClip1.GraphicCell(y)
If Me.Picture <> 0 Then
  Call SetAutoRgn(Me)
End If
End Sub
