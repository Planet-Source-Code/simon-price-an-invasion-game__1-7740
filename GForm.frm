VERSION 5.00
Begin VB.Form GForm 
   Caption         =   "Invasion - By Simon Price"
   ClientHeight    =   5352
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   7176
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   446
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   598
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer AddMoreMayhem 
      Interval        =   4000
      Left            =   960
      Top             =   2880
   End
   Begin VB.PictureBox UFOPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawWidth       =   10
      Height          =   360
      Index           =   5
      Left            =   3240
      Picture         =   "GForm.frx":0000
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox UFOPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawWidth       =   10
      Height          =   360
      Index           =   4
      Left            =   3240
      Picture         =   "GForm.frx":0A5A
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox UFOPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawWidth       =   10
      Height          =   360
      Index           =   3
      Left            =   3240
      Picture         =   "GForm.frx":14B4
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox UFOPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawWidth       =   10
      Height          =   360
      Index           =   2
      Left            =   3240
      Picture         =   "GForm.frx":1F0E
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox UFOPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawWidth       =   10
      Height          =   360
      Index           =   1
      Left            =   3240
      Picture         =   "GForm.frx":2968
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox UFOMask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawWidth       =   10
      Height          =   360
      Left            =   4200
      Picture         =   "GForm.frx":33C2
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox BulletMask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawWidth       =   10
      Height          =   120
      Left            =   4920
      Picture         =   "GForm.frx":3E1C
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox BulletPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawWidth       =   10
      Height          =   120
      Index           =   5
      Left            =   4200
      Picture         =   "GForm.frx":42D6
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox BulletPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawWidth       =   10
      Height          =   120
      Index           =   4
      Left            =   3960
      Picture         =   "GForm.frx":4790
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox BulletPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawWidth       =   10
      Height          =   120
      Index           =   3
      Left            =   3720
      Picture         =   "GForm.frx":4C4A
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox BulletPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawWidth       =   10
      Height          =   120
      Index           =   2
      Left            =   3480
      Picture         =   "GForm.frx":5104
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox BulletPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawWidth       =   10
      Height          =   120
      Index           =   1
      Left            =   3240
      Picture         =   "GForm.frx":55BE
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox PB 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawWidth       =   10
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H000000FF&
      Height          =   2292
      Left            =   120
      ScaleHeight     =   191
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2412
   End
End
Attribute VB_Name = "GForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DispWidth As Integer
Private DispHeight As Integer
Private ExitLoop As Boolean

Private Sub AddMoreMayhem_Timer()
i = UBound(Invader) + 1
ReDim Preserve Invader(1 To i)
  Invader(i).X = (Rnd * 500) + 50
  Invader(i).Y = -(Rnd * 10)
  Invader(i).xm = Int(Rnd * 10) - 5
  Invader(i).ym = 5
  Invader(i).Frame = Int(Rnd * 5) + 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyEscape
    ExitLoop = True
    SForm.Visible = True
    Unload Me
  Case vbKeyLeft
    Shooter.Angle = Shooter.Angle + 10
  Case vbKeyRight
    Shooter.Angle = Shooter.Angle - 10
  Case vbKeySpace
    If BulletNo = BULLETS Then
      BulletNo = 1
    Else
    BulletNo = BulletNo + 1
    End If
        Bullet(BulletNo).X = 300
        Bullet(BulletNo).Y = 450
        Bullet(BulletNo).xm = -20 * Sin(Shooter.Angle * PIdiv180)
        Bullet(BulletNo).ym = -20 * Cos(Shooter.Angle * PIdiv180)
  Case vbKeyC
    PB.Picture = PB.Image
    SavePicture PB.Picture, App.Path & "/Pic.bmp"
    PB.Picture = LoadPicture()
End Select
End Sub

Private Sub Form_Load()
Randomize Timer
SortLayout
Show
CreateStars
CreateInvaders
MainLoop
End Sub

Public Sub MainLoop()
Do
DoEvents
PB.Cls
If StarsOn Then DoStars
DoShooter
DoInvaders
DoBullets
DoExplosions
StretchBlt hdc, 0, 0, DispWidth, DispHeight, PB.hdc, 0, 0, 600, 450, vbSrcCopy
Loop Until ExitLoop = True
End Sub

Private Sub SortLayout()
Move 0, 0, Screen.Width, Screen.Height
PB.Move 0, 0, 600, 450
DispWidth = Screen.Height * 1.315 / Screen.TwipsPerPixelY
DispHeight = Screen.Height * 0.975 / Screen.TwipsPerPixelY
ScaleWidth = 180
ScaleHeight = 100
End Sub

Private Sub CreateStars()
For i = 1 To WHITESTARS
  WhiteStar(i).X = Rnd * 600
  WhiteStar(i).Y = Rnd * 450
Next
For i = 1 To CYANSTARS
  CyanStar(i).X = Rnd * 600
  CyanStar(i).Y = Rnd * 450
Next
For i = 1 To YELLOWSTARS
  YellowStar(i).X = Rnd * 600
  YellowStar(i).Y = Rnd * 450
Next
For i = 1 To REDSTARS
  RedStar(i).X = Rnd * 600
  RedStar(i).Y = Rnd * 450
Next
End Sub

Private Sub CreateInvaders()
Difficulty = 5
ReDim Invader(1 To 5)
For i = 1 To 5
  Invader(i).X = (Rnd * 500) + 50
  Invader(i).Y = -(Rnd * 10)
  Invader(i).xm = Int(Rnd * 10) - 5
  Invader(i).ym = 5
  Invader(i).Frame = Int(Rnd * 5) + 1
Next
End Sub

Private Sub DoExplosions()
PB.DrawWidth = 1
For i = 1 To EXPLOSIONS
If Explosion(i).Power <> 0 Then
    Explosion(i).Power = Explosion(i).Power - 1
    For i2 = 1 To Explosion(i).Power
      MoveToEx PB.hdc, Explosion(i).Pixel(i2).X, Explosion(i).Pixel(i2).Y, lpPoint
      Explosion(i).Pixel(i2).X = Explosion(i).Pixel(i2).X + Explosion(i).Pixel(i2).xm
      Explosion(i).Pixel(i2).Y = Explosion(i).Pixel(i2).Y + Explosion(i).Pixel(i2).ym
      PB.ForeColor = Explosion(i).Pixel(i2).Color
      LineTo PB.hdc, Explosion(i).Pixel(i2).X, Explosion(i).Pixel(i2).Y
    Next
End If
Next
End Sub

Private Sub CreateExplosion()
'uncomment this line below for explosion sounds!!!
'sndPlaySound App.Path & "\Explode.wav", &H1
If ExplosionNo = EXPLOSIONS Then
  ExplosionNo = 1
Else
  ExplosionNo = ExplosionNo + 1
End If
Explosion(ExplosionNo).Power = EXPLOSIONPOWER
For i3 = 1 To EXPLOSIONPOWER
  Explosion(ExplosionNo).Pixel(i3).X = X
  Explosion(ExplosionNo).Pixel(i3).Y = Y
  Explosion(ExplosionNo).Pixel(i3).xm = (Rnd * 20) - 10
  Explosion(ExplosionNo).Pixel(i3).ym = (Rnd * 20) - 10
    Select Case Int(Rnd * 4)
      Case 0: Explosion(ExplosionNo).Pixel(i3).Color = vbWhite
      Case 1: Explosion(ExplosionNo).Pixel(i3).Color = vbYellow
      Case 2: Explosion(ExplosionNo).Pixel(i3).Color = vbCyan
      Case 3: Explosion(ExplosionNo).Pixel(i3).Color = vbRed
    End Select
Next
End Sub

Private Sub DoInvaders()
For i = 1 To UBound(Invader)
    If Invader(i).Frame = 5 Then
      Invader(i).Frame = 1
    Else
      Invader(i).Frame = Invader(i).Frame + 1
    End If
  X = Invader(i).X + Invader(i).xm
  Y = Invader(i).Y + Invader(i).ym
    Select Case X
      Case Is < 0
        Invader(i).xm = -Invader(i).xm
      Case Is > 600
        Invader(i).xm = -Invader(i).xm
      Case Else
        Invader(i).X = X
    End Select
    Select Case Y
      Case Is > 450
        MsgBox "Game Over - the invaders have invaded!!!", vbExclamation, "U IS DEAD!"
        ExitLoop = True
        End
      Case Else
        Invader(i).Y = Y
    End Select
        BitBlt PB.hdc, Invader(i).X - 25, Invader(i).Y - 15, 50, 30, UFOMask.hdc, 0, 0, vbMergePaint
        BitBlt PB.hdc, Invader(i).X - 25, Invader(i).Y - 15, 50, 30, UFOPic(Invader(i).Frame).hdc, 0, 0, vbSrcAnd
Next
End Sub

Private Sub DoShooter()
PB.DrawWidth = 10
PB.ForeColor = vbRed
MoveToEx PB.hdc, 300, 450, lpPoint
SinA = Sin(Shooter.Angle * PIdiv180)
CosA = Cos(Shooter.Angle * PIdiv180)
LineTo PB.hdc, 300 - 50 * SinA, 450 - 50 * CosA
End Sub

Private Sub DoBullets()
For i = 1 To BULLETS
  If Bullet(i).X > 0 Then
    X = Bullet(i).X + Bullet(i).xm
    Y = Bullet(i).Y + Bullet(i).ym
  For i2 = 1 To UBound(Invader)
    Select Case Invader(i2).X
      Case X - 30 To X + 30
        Select Case Invader(i2).Y
          Case Y - 20 To Y + 20
            Invader(i2).X = (Rnd * 500) + 50
            Invader(i2).Y = (Rnd * 100) - 100
            Invader(i2).xm = Int(Rnd * 10) - 5
            Invader(i2).ym = Difficulty
            Shooter.Score = Shooter.Score + Invader(i2).Y
            Caption = "Invasion - By Simon Price - Score = " & -Shooter.Score & " points"
            CreateExplosion
        End Select
    End Select
  Next
  Bullet(i).X = X
  Bullet(i).Y = Y
    BitBlt PB.hdc, Bullet(i).X - 5, Bullet(i).Y - 5, 10, 10, BulletMask.hdc, 0, 0, vbMergePaint
    BitBlt PB.hdc, Bullet(i).X - 5, Bullet(i).Y - 5, 10, 10, BulletPic((i Mod 5) + 1).hdc, 0, 0, vbSrcAnd
End If
Next
End Sub

Private Sub DoStars()
'white stars
For i = 1 To WHITESTARS
  If WhiteStar(i).Y > 600 Then
    WhiteStar(i).X = Rnd * 600
    WhiteStar(i).Y = 0
  Else
  WhiteStar(i).Y = WhiteStar(i).Y + 2
  SetPixel PB.hdc, WhiteStar(i).X, WhiteStar(i).Y, vbWhite
  End If
Next
'cyan stars
For i = 1 To CYANSTARS
  If CyanStar(i).Y > 600 Then
    CyanStar(i).X = Rnd * 600
    CyanStar(i).Y = 0
  Else
  CyanStar(i).Y = CyanStar(i).Y + 3
  SetPixel PB.hdc, CyanStar(i).X, CyanStar(i).Y, vbCyan
  End If
Next
'yellow stars
For i = 1 To YELLOWSTARS
  If YellowStar(i).Y > 600 Then
    YellowStar(i).X = Rnd * 600
    YellowStar(i).Y = 0
  Else
  YellowStar(i).Y = YellowStar(i).Y + 5
  SetPixel PB.hdc, YellowStar(i).X, YellowStar(i).Y, vbYellow
  End If
Next
'red stars
For i = 1 To REDSTARS
  If RedStar(i).Y > 600 Then
    RedStar(i).X = Rnd * 600
    RedStar(i).Y = 0
  Else
  RedStar(i).Y = RedStar(i).Y + 8
  SetPixel PB.hdc, RedStar(i).X, RedStar(i).Y, vbRed
  End If
Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shooter.Angle = 90 - X
End Sub
