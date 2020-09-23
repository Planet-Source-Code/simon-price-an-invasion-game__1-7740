Attribute VB_Name = "InvadersMod"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_ASYNC = &H1         '  play asynchronously

Public Type POINTAPI
  X As Long
  Y As Long
End Type

Public Type tStar
  X As Integer
  Y As Integer
End Type

Public Type tInvader
  X As Integer
  Y As Integer
  xm As Integer
  ym As Integer
  Frame As Byte
End Type

Public Type tShooter
  Angle As Integer
  Score As Long
End Type

Public Type tExplodePiece
  X As Integer
  Y As Integer
  xm As Integer
  ym As Integer
  Color As ColorConstants
End Type

Public EXPLOSIONPOWER As Byte
Public StarsOn As Boolean

Public Type tExplosion
  Pixel() As tExplodePiece
  Power As Byte
End Type

Public Type tBullet
  X As Integer
  Y As Integer
  xm As Integer
  ym As Integer
End Type

Public Invader() As tInvader
Public Shooter As tShooter
Public Const EXPLOSIONS = 10
Public Explosion(1 To EXPLOSIONS) As tExplosion
Public ExplosionNo As Byte
Public Const BULLETS = 50
Public Bullet(1 To BULLETS) As tBullet
Public BulletNo As Byte

Public Const WHITESTARS = 50
Public Const CYANSTARS = 25
Public Const YELLOWSTARS = 10
Public Const REDSTARS = 5
Public WhiteStar(1 To WHITESTARS) As tStar
Public CyanStar(1 To CYANSTARS) As tStar
Public YellowStar(1 To YELLOWSTARS) As tStar
Public RedStar(1 To REDSTARS) As tStar

Public i As Integer
Public i2 As Integer
Public i3 As Integer
Public X As Integer
Public Y As Integer

Public Difficulty As Byte

Public Const PI = 3.14159265358979
Public Const PIdiv180 = PI / 180

Public lpPoint As POINTAPI
Public SinA As Single
Public CosA As Single

