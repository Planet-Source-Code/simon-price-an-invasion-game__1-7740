VERSION 5.00
Begin VB.Form SForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invasion - By Simon Price - www.hispalace.fsbusiness.co.uk !!!!!"
   ClientHeight    =   4788
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   6336
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4788
   ScaleWidth      =   6336
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO !!!"
      Height          =   492
      Left            =   2640
      TabIndex        =   6
      Top             =   2160
      Width           =   1332
   End
   Begin VB.Frame Frame1 
      Caption         =   "Graphics Setup"
      Height          =   4332
      Left            =   4200
      TabIndex        =   1
      Top             =   240
      Width           =   1932
      Begin VB.HScrollBar ExplosionO 
         Height          =   252
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   3
         Top             =   2280
         Value           =   40
         Width           =   1692
      End
      Begin VB.CheckBox StarsO 
         BackColor       =   &H0000FF00&
         Caption         =   "Draw background (there are stars that you appear to be flying over in the bakground)"
         Height          =   972
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1692
      End
      Begin VB.Label ExplosionL 
         BackColor       =   &H0000FF00&
         Caption         =   "Explosion Power = 40"
         Height          =   252
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   1692
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FF00&
         Caption         =   "Change these options if the graphics are too slow, but try defaults 1st."
         Height          =   612
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1692
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   1800
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   $"SForm.frx":0000
         Height          =   1332
         Left            =   120
         TabIndex        =   4
         Top             =   2880
         Width           =   1692
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   1800
         Y1              =   2160
         Y2              =   2160
      End
   End
   Begin VB.Label InfoL 
      BackColor       =   &H0000FF00&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0000FF00&
      Height          =   4332
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2292
   End
   Begin VB.Image Pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   4572
      Left            =   120
      Picture         =   "SForm.frx":009D
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6132
   End
End
Attribute VB_Name = "SForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGo_Click()
EXPLOSIONPOWER = ExplosionO.Value
For i = 1 To 10
  ReDim Explosion(i).Pixel(1 To EXPLOSIONPOWER)
Next
StarsOn = StarsO.Value
MsgBox "Use the mouse (MUST BE ON FORM!) to control your aim. Press the spacebar to shoot. Press Escape to return to this screen if you need to change the graphics settings.", vbInformation, "Controls :"
Hide
GForm.Visible = True
Unload Me
End Sub

Private Sub ExplosionO_Change()
ExplosionL = "Explosion Power = " & ExplosionO.Value
End Sub

Private Sub Form_Load()
Dim Message As String
Message = "This game is demo about using the API's" & _
         " BitBlt, SetPixel, MoveToEx and LineTo" & _
         " to make graphics that are much faster" & _
         " than VB's slower methods. I use BitBlt and" & _
         " masking to draw UFO's, I use SetPixel to create" & _
         " stars, and MoveToEx/LineTo to create a" & _
         " cool explosion effect. The graphics are very" & _
         " detailed and fast on my computer, however, mine is 400Mhz" & _
         " and so I expect that this game will noy perform as well" & _
         " on slower machines. This is why I have given you a few" & _
         " options to trade graphics for speed. I have it set with all" & _
         " options at what works for me and I recommend you try that first" & _
         " and only start changing the options if you need to."
InfoL = Message
End Sub
