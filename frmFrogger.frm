VERSION 5.00
Begin VB.Form frmFrogger 
   BackColor       =   &H00000000&
   Caption         =   "FROGGER"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9015
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   9015
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer snakeTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3120
      Top             =   3480
   End
   Begin VB.Timer collisionTimer 
      Interval        =   20
      Left            =   2520
      Top             =   3480
   End
   Begin VB.Timer scoreTimer 
      Interval        =   100
      Left            =   1920
      Top             =   3480
   End
   Begin VB.Timer gameTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1320
      Top             =   3480
   End
   Begin VB.Timer riverTimer 
      Interval        =   80
      Left            =   720
      Top             =   3480
   End
   Begin VB.Timer roadTimer 
      Interval        =   80
      Left            =   120
      Top             =   3480
   End
   Begin VB.Image imgSnake 
      Height          =   360
      Left            =   3720
      Picture         =   "frmFrogger.frx":0000
      Top             =   3480
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label lblSplash 
      BackColor       =   &H00004000&
      Caption         =   $"frmFrogger.frx":0139
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   9015
   End
   Begin VB.Label lblLevel 
      BackColor       =   &H00008000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8160
      TabIndex        =   5
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Caption         =   "LEVEL:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
   Begin VB.Image imgCar 
      Height          =   600
      Index           =   10
      Left            =   5880
      Picture         =   "frmFrogger.frx":023A
      Top             =   5160
      Width           =   600
   End
   Begin VB.Image imgCar 
      Height          =   600
      Index           =   9
      Left            =   2880
      Picture         =   "frmFrogger.frx":03A1
      Top             =   5160
      Width           =   1800
   End
   Begin VB.Image imgCar 
      Height          =   600
      Index           =   11
      Left            =   7800
      Picture         =   "frmFrogger.frx":051E
      Top             =   5160
      Width           =   1200
   End
   Begin VB.Image imgCar 
      Height          =   600
      Index           =   7
      Left            =   7080
      Picture         =   "frmFrogger.frx":0744
      Top             =   4560
      Width           =   1200
   End
   Begin VB.Image imgCar 
      Height          =   600
      Index           =   3
      Left            =   7440
      Picture         =   "frmFrogger.frx":096A
      Top             =   3960
      Width           =   600
   End
   Begin VB.Image imgCar 
      Height          =   600
      Index           =   2
      Left            =   5760
      Picture         =   "frmFrogger.frx":0AD4
      Top             =   3960
      Width           =   600
   End
   Begin VB.Image imgCar 
      Height          =   600
      Index           =   6
      Left            =   4560
      Picture         =   "frmFrogger.frx":0C3E
      Top             =   4560
      Width           =   1200
   End
   Begin VB.Image imgCar 
      Height          =   600
      Index           =   5
      Left            =   2640
      Picture         =   "frmFrogger.frx":0E64
      Top             =   4560
      Width           =   600
   End
   Begin VB.Image imgCar 
      Height          =   600
      Index           =   1
      Left            =   3120
      Picture         =   "frmFrogger.frx":0FCB
      Top             =   3960
      Width           =   1200
   End
   Begin VB.Image imgCar 
      Height          =   600
      Index           =   8
      Left            =   960
      Picture         =   "frmFrogger.frx":11F5
      Top             =   5160
      Width           =   600
   End
   Begin VB.Image imgCar 
      Height          =   600
      Index           =   4
      Left            =   600
      Picture         =   "frmFrogger.frx":135C
      Top             =   4560
      Width           =   1200
   End
   Begin VB.Image imgCar 
      Height          =   600
      Index           =   0
      Left            =   240
      Picture         =   "frmFrogger.frx":1582
      Top             =   3960
      Width           =   1800
   End
   Begin VB.Image imgFrog 
      Height          =   600
      Index           =   4
      Left            =   7800
      Picture         =   "frmFrogger.frx":1701
      Top             =   960
      Width           =   600
   End
   Begin VB.Image imgFrog 
      Height          =   600
      Index           =   3
      Left            =   6000
      Picture         =   "frmFrogger.frx":1880
      Top             =   960
      Width           =   600
   End
   Begin VB.Image imgFrog 
      Height          =   600
      Index           =   2
      Left            =   4200
      Picture         =   "frmFrogger.frx":19FF
      Top             =   960
      Width           =   600
   End
   Begin VB.Image imgFrog 
      Height          =   600
      Index           =   1
      Left            =   2400
      Picture         =   "frmFrogger.frx":1B7E
      Top             =   960
      Width           =   600
   End
   Begin VB.Image imgFrog 
      Height          =   600
      Index           =   0
      Left            =   600
      Picture         =   "frmFrogger.frx":1CFD
      Top             =   960
      Width           =   600
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   9000
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   9000
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   2  'Dash
      X1              =   0
      X2              =   9000
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FFFF&
      X1              =   0
      X2              =   9000
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FFFF&
      X1              =   0
      X2              =   9000
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      X1              =   0
      X2              =   9000
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      X1              =   0
      X2              =   9000
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Image imgLog 
      Height          =   600
      Index           =   11
      Left            =   6120
      Picture         =   "frmFrogger.frx":1E7C
      Top             =   2760
      Width           =   1800
   End
   Begin VB.Image imgLog 
      Height          =   600
      Index           =   10
      Left            =   3600
      Picture         =   "frmFrogger.frx":2055
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Image imgLog 
      Height          =   600
      Index           =   9
      Left            =   240
      Picture         =   "frmFrogger.frx":21C2
      Top             =   2760
      Width           =   1800
   End
   Begin VB.Image imgLog 
      Height          =   600
      Index           =   8
      Left            =   6840
      Picture         =   "frmFrogger.frx":239B
      Top             =   2160
      Width           =   1800
   End
   Begin VB.Image imgLog 
      Height          =   600
      Index           =   7
      Left            =   4920
      Picture         =   "frmFrogger.frx":2574
      Top             =   2160
      Width           =   1200
   End
   Begin VB.Image imgLog 
      Height          =   600
      Index           =   6
      Left            =   2160
      Picture         =   "frmFrogger.frx":26E1
      Top             =   2160
      Width           =   1800
   End
   Begin VB.Image imgLog 
      Height          =   600
      Index           =   5
      Left            =   480
      Picture         =   "frmFrogger.frx":28BA
      Top             =   2160
      Width           =   600
   End
   Begin VB.Image imgLog 
      Height          =   600
      Index           =   4
      Left            =   7680
      Picture         =   "frmFrogger.frx":29B1
      Top             =   1560
      Width           =   1200
   End
   Begin VB.Image imgLog 
      Height          =   600
      Index           =   2
      Left            =   4560
      Picture         =   "frmFrogger.frx":2B1E
      Top             =   1560
      Width           =   600
   End
   Begin VB.Image imgLog 
      Height          =   600
      Index           =   3
      Left            =   5880
      Picture         =   "frmFrogger.frx":2C15
      Top             =   1560
      Width           =   1200
   End
   Begin VB.Image imgLog 
      Height          =   600
      Index           =   1
      Left            =   1920
      Picture         =   "frmFrogger.frx":2D82
      Top             =   1560
      Width           =   1800
   End
   Begin VB.Image imgLog 
      Height          =   600
      Index           =   0
      Left            =   120
      Picture         =   "frmFrogger.frx":2F5B
      Top             =   1560
      Width           =   1200
   End
   Begin VB.Image imgLife 
      Height          =   375
      Index           =   4
      Left            =   6600
      Picture         =   "frmFrogger.frx":30C8
      Stretch         =   -1  'True
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgLife 
      Height          =   375
      Index           =   3
      Left            =   7080
      Picture         =   "frmFrogger.frx":31E5
      Stretch         =   -1  'True
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgLife 
      Height          =   375
      Index           =   2
      Left            =   7560
      Picture         =   "frmFrogger.frx":3302
      Stretch         =   -1  'True
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgLife 
      Height          =   375
      Index           =   1
      Left            =   8040
      Picture         =   "frmFrogger.frx":341F
      Stretch         =   -1  'True
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgLife 
      Height          =   375
      Index           =   0
      Left            =   8520
      Picture         =   "frmFrogger.frx":353C
      Stretch         =   -1  'True
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblHighScore 
      BackColor       =   &H00008000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label lblHScore 
      BackColor       =   &H00008000&
      Caption         =   "HIGHSCORE:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00008000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblScor 
      BackColor       =   &H00008000&
      Caption         =   "SCORE:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Shape shpBottomGrass 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   0
      Top             =   5760
      Width           =   9015
   End
   Begin VB.Shape shpRoad 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   0
      Top             =   3960
      Width           =   9015
   End
   Begin VB.Shape shpCentreGrass 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   3360
      Width           =   9015
   End
   Begin VB.Shape shpRiver 
      BackColor       =   &H00C0C000&
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   0
      Top             =   1560
      Width           =   9015
   End
   Begin VB.Image Image5 
      Height          =   1095
      Left            =   7200
      Picture         =   "frmFrogger.frx":3659
      Top             =   480
      Width           =   1815
   End
   Begin VB.Image Image4 
      Height          =   1095
      Left            =   5400
      Picture         =   "frmFrogger.frx":9E67
      Top             =   480
      Width           =   1815
   End
   Begin VB.Image Image3 
      Height          =   1095
      Left            =   1800
      Picture         =   "frmFrogger.frx":10675
      Top             =   480
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   0
      Picture         =   "frmFrogger.frx":16E83
      Top             =   480
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   3600
      Picture         =   "frmFrogger.frx":1D691
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmFrogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim m As Integer
Dim frogIndex As Integer
Dim lives As Integer
Dim score As Long
Dim highScore As Long
Dim frogState(4) As Integer
Dim home(4) As Boolean
Dim hassound As Integer, retval As Integer
Dim logSpeed(11) As Integer
Dim numOfLogs As Integer
Dim extraLifeCount As Integer
Dim frogLog(11) As Boolean
Dim dry As Boolean
Dim carSpeed(11) As Integer
Dim numOfCars As Integer
Dim level As Integer
Dim result  As Variant
Dim snakeAlive As Boolean
Dim swap As Integer
Dim swapCount As Integer

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
     
Private Declare Function PlaySound Lib "winmm.dll" Alias _
    "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, _
        ByVal dwFlags As Long) As Long
Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal LpstrCommand As String, ByVal LpstrReturnString As String, ByVal UReturnLength As Long, ByVal HwndCallBack As Long) As Long

Public Function PlayWaveFile(strFileName As String, Optional blnAsync As Boolean) As Boolean
    Dim lngFlags As Long
    Const snd_sync = &H0
    Const snd_Async = &H1
    Const snd_Nodefault = &H2
    Const snd_Filename = &H20000
    lngFlags = snd_Nodefault Or snd_Filename Or snd_sync
    If blnAsync Then lngFlags = lngFlags Or snd_Async
    PlayWaveFile = PlaySound(strFileName, 0&, lngFlags)
End Function

Private Sub collisionTimer_Timer()

k = imgFrog(frogIndex).Top
If (k <> 2760 And k <> 2160 And k <> 1560) Then
  GoTo 300:
Else
  'in river
  
  For i = 0 To numOfLogs - 1
    If frogLog(i) = True Then GoTo 600
  Next i
  For i = 0 To numOfLogs - 1
    If imgLog(i).Visible = True And (imgLog(i).Left - 250 <= imgFrog(frogIndex).Left) _
    And (imgFrog(frogIndex).Left + 600 <= imgLog(i).Left + imgLog(i).Width + 250) _
    And (imgLog(i).Top = k) Then
      For j = 0 To numOfLogs - 1
        frogLog(j) = False
      Next j
      frogLog(i) = True
      GoTo 600
    End If
  Next i
600:
dry = False
  For i = 0 To 11
    If frogLog(i) = True Then dry = True
  Next i
  If dry = False Then die
End If
300:
If (k <> 5160 And k <> 4560 And k <> 3960) Then
  GoTo 400
Else
  'in road
  For i = 0 To numOfCars - 1
    If imgCar(i).Top = imgFrog(frogIndex).Top _
    And imgCar(i).Visible = True _
    And (imgCar(i).Left + imgCar(i).Width >= imgFrog(frogIndex).Left And imgCar(i).Left <= imgFrog(frogIndex).Left + 600) _
    Then die
  Next i
End If
400:
If (k <> 3360) Then
  GoTo 500
Else
  'on grass, (snake)
  If imgSnake.Left < imgFrog(frogIndex).Left + imgFrog(frogIndex).Width _
  And imgSnake.Left + imgSnake.Width > imgFrog(frogIndex).Left Then die
End If
500:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then moveUp
If KeyCode = 40 Then moveDown
If KeyCode = 37 Then moveLeft
If KeyCode = 39 Then moveRight
If KeyCode = 27 Or KeyCode = 81 Then End
If KeyCode = 32 Then initialiseGame
End Sub
Private Sub moveUp()
For m = 0 To numOfLogs - 1
  frogLog(m) = False
Next m
If imgFrog(frogIndex).Top > 1560 Then
  'change image
  imgFrog(frogIndex) = LoadPicture(App.Path & "\" & "frogU2.gif")
  'play sound
  retval = PlayWaveFile(App.Path & "\" & "Hop.wav", True)
  score = score + 100
  imgFrog(frogIndex).Top = imgFrog(frogIndex).Top - 600
  imgFrog(frogIndex) = LoadPicture(App.Path & "\" & "frogU1.gif")
Else
  'home check
  homeCheck
End If
DoEvents
End Sub
Private Sub moveDown()
For m = 0 To numOfLogs - 1
  frogLog(m) = False
Next m
If imgFrog(frogIndex).Top < 5760 Then
  'change image
  imgFrog(frogIndex) = LoadPicture(App.Path & "\" & "frogD2.gif")
  'play sound
  retval = PlayWaveFile(App.Path & "\" & "Hop.wav", True)
  score = score + 50
  imgFrog(frogIndex).Top = imgFrog(frogIndex).Top + 600
  imgFrog(frogIndex) = LoadPicture(App.Path & "\" & "frogD1.gif")
End If
DoEvents
End Sub
Private Sub moveLeft()
For m = 0 To numOfLogs - 1
  frogLog(m) = False
Next m
If imgFrog(frogIndex).Left > 0 Then
  'change image
  imgFrog(frogIndex) = LoadPicture(App.Path & "\" & "frogL2.gif")
  'play sound
  retval = PlayWaveFile(App.Path & "\" & "Hop.wav", True)
  score = score + 1
  imgFrog(frogIndex).Left = imgFrog(frogIndex).Left - 600
  imgFrog(frogIndex) = LoadPicture(App.Path & "\" & "frogL1.gif")
End If
DoEvents
End Sub
Private Sub moveRight()
For m = 0 To numOfLogs - 1
  frogLog(m) = False
Next m
If imgFrog(frogIndex).Left < 8400 Then
  'change image
  imgFrog(frogIndex) = LoadPicture(App.Path & "\" & "frogR2.gif")
  'play sound
  retval = PlayWaveFile(App.Path & "\" & "Hop.wav", True)
  score = score + 1
  imgFrog(frogIndex).Left = imgFrog(frogIndex).Left + 600
  imgFrog(frogIndex) = LoadPicture(App.Path & "\" & "frogR1.gif")
End If
DoEvents
End Sub
Private Sub Form_Load()

initialiseGame
snakeTimer.Enabled = False
scoreTimer.Enabled = False
roadTimer.Enabled = False
riverTimer.Enabled = False
collisionTimer.Enabled = False
gameTimer.Enabled = False
splash

End Sub

Private Sub initialiseGame()
lives = 4
score = 0
level = 0
numOfLogs = 12
numOfCars = 12
For i = 0 To 4: logSpeed(i) = 75: Next i
For i = 5 To 8: logSpeed(i) = -50: Next i
For i = 9 To 11: logSpeed(i) = 25: Next i
For i = 0 To 3: carSpeed(i) = 50: Next i
For i = 4 To 7: carSpeed(i) = -50: Next i
For i = 8 To 11: carSpeed(i) = -25: Next i
For i = 0 To 4
  frogState(i) = 0
  imgFrog(i).Visible = False
Next i
For i = 0 To 4
  home(i) = False
Next i
imgCar(3).Visible = False
imgCar(6).Visible = False
imgCar(10).Visible = False
imgCar(4).Visible = False
imgCar(0).Visible = False
imgLog(3).Visible = True
imgLog(6).Visible = True
imgLog(11).Visible = True
snakeTimer.Enabled = True
scoreTimer.Enabled = True
roadTimer.Enabled = True
riverTimer.Enabled = True
collisionTimer.Enabled = True
gameTimer.Enabled = True
For i = 0 To 4
  imgLife(i).Visible = False
Next i
If lives - 1 > 4 Then j = 4 Else j = lives - 1
For i = 0 To j
  imgLife(i).Visible = True
Next i
'start music
newFrog
End Sub

Private Sub gameEnd()
retval = PlayWaveFile(App.Path & "\" & "Gameover.wav", True)
Sleep (3000)
'gameover, highscore
'play again?
snakeTimer.Enabled = False
scoreTimer.Enabled = False
roadTimer.Enabled = False
riverTimer.Enabled = False
collisionTimer.Enabled = False
gameTimer.Enabled = False
splash
End Sub
Private Sub splash()
lblSplash.Visible = True
result = MsgBox("Start Game?")
lblSplash.Visible = False
initialiseGame
End Sub
Private Sub newFrog()
'select frog from 0-4
  While frogState(frogIndex) <> 0
    frogIndex = Int(Rnd * 5)
  Wend
  imgFrog(frogIndex).Visible = True
  imgFrog(frogIndex).Top = 5760
  imgFrog(frogIndex).Left = 4200
  numOfLogs = 12 '?
  For i = 0 To numOfLogs - 1
    frogLog(i) = False
  Next i
End Sub
Private Sub gameTimer_Timer()
If lives >= 1 Then
  If frogState(0) = 1 And frogState(1) = 1 And frogState(2) = 1 And frogState(3) = 1 And frogState(4) = 1 Then
    'all frogs home
    'reset frogs states
    For i = 0 To 4
      frogState(i) = 0
    Next i
    newFrog
  End If
Else
  gameEnd
End If
End Sub

Private Sub riverTimer_Timer()
'river
For l = 0 To numOfLogs - 1
  If frogLog(l) = True Then
    imgFrog(frogIndex).Left = imgFrog(frogIndex).Left + logSpeed(l)
    If (imgFrog(frogIndex).Left) < 0 Or imgFrog(frogIndex).Left + imgFrog(frogIndex).Width > frmFrogger.ScaleWidth Then die
  End If
  imgLog(l).Left = imgLog(l).Left + logSpeed(l)
  If imgLog(l).Left > frmFrogger.ScaleWidth + (200 * Int(Rnd * 3)) Then imgLog(l).Left = 0 - imgLog(l).Width: GoTo 200
  If imgLog(l).Left < (0 - imgLog(l).Width - (200 * Int(Rnd * 2))) Then imgLog(l).Left = frmFrogger.ScaleWidth
200:
Next l

End Sub

Private Sub roadTimer_Timer()
'road
For l = 0 To numOfCars - 1
  imgCar(l).Left = imgCar(l).Left + carSpeed(l)
  If imgCar(l).Left > frmFrogger.ScaleWidth + (50) Then imgCar(l).Left = 0 - imgCar(l).Width: GoTo 1000
  If imgCar(l).Left < (0 - imgCar(l).Width - (50)) Then imgCar(l).Left = frmFrogger.ScaleWidth
1000:
Next l
End Sub

Private Sub homeCheck()

j = imgFrog(frogIndex).Left
If (j > 500 And j < 700 And home(0) = False) Then
  home(0) = True
  imgFrog(frogIndex).Top = 960
  imgFrog(frogIndex).Left = 600
  'load picture
  imgFrog(frogIndex) = LoadPicture(App.Path & "\" & "frogHome.gif")
  Refresh
  frogState(frogIndex) = 1
  allFrogsHomeCheck
  newFrog
Else
  If (j > 2300 And j < 2500 And home(1) = False) Then
    home(1) = True
    imgFrog(frogIndex).Top = 960
    imgFrog(frogIndex).Left = 2400
    'load picture
    imgFrog(frogIndex) = LoadPicture(App.Path & "\" & "frogHome.gif")
    Refresh
    frogState(frogIndex) = 1
    allFrogsHomeCheck
    newFrog
  Else
    If (j > 4100 And j < 4300 And home(2) = False) Then
      home(2) = True
      imgFrog(frogIndex).Top = 960
      imgFrog(frogIndex).Left = 4200
      'load picture
      imgFrog(frogIndex) = LoadPicture(App.Path & "\" & "frogHome.gif")
      Refresh
      frogState(frogIndex) = 1
      allFrogsHomeCheck
      newFrog
    Else
      If (j > 5900 And j < 6100 And home(3) = False) Then
        home(3) = True
        imgFrog(frogIndex).Top = 960
        imgFrog(frogIndex).Left = 6000
        'load picture
        imgFrog(frogIndex) = LoadPicture(App.Path & "\" & "frogHome.gif")
        Refresh
        frogState(frogIndex) = 1
        allFrogsHomeCheck
        newFrog
      Else
        If (j > 7700 And j < 7900 And home(4) = False) Then
          home(4) = True
          imgFrog(frogIndex).Top = 960
          imgFrog(frogIndex).Left = 7800
          'load picture
          imgFrog(frogIndex) = LoadPicture(App.Path & "\" & "frogHome.gif")
          Refresh
          frogState(frogIndex) = 1
          allFrogsHomeCheck
          newFrog
        Else
          'die
          die
        End If
      End If
    End If
  End If
End If
End Sub

Private Sub die()
retval = PlayWaveFile(App.Path & "\" & "Die.wav", True)
For i = 1 To 3
  imgFrog(frogIndex) = LoadPicture(App.Path & "\" & "frogU1.gif")
  imgFrog(frogIndex).Refresh
  Sleep (100)
  imgFrog(frogIndex) = LoadPicture(App.Path & "\" & "frogR1.gif")
  imgFrog(frogIndex).Refresh
  Sleep (100)
  imgFrog(frogIndex) = LoadPicture(App.Path & "\" & "frogD1.gif")
  imgFrog(frogIndex).Refresh
  Sleep (100)
  imgFrog(frogIndex) = LoadPicture(App.Path & "\" & "frogL1.gif")
  imgFrog(frogIndex).Refresh
  Sleep (100)
  imgFrog(frogIndex) = LoadPicture(App.Path & "\" & "frogU1.gif")
  imgFrog(frogIndex).Refresh
  Sleep (100)
Next i
Sleep (500)
lives = lives - 1
If lives = 0 Then
  gameEnd
Else
  
  newFrog
End If
End Sub
Private Sub allFrogsHomeCheck()
retval = PlayWaveFile(App.Path & "\" & "Home.wav", True)
Sleep (1500)
j = 0
For i = 0 To 4
  If home(i) = True Then j = j + 1
Next i
If j = 5 Then newLevel
End Sub

Private Sub newLevel()
'score bonus
score = score + 1000
level = level + 1
extraLifeCount = extraLifeCount + 1
If extraLifeCount = 3 Then extraLifeCount = 0: lives = lives + 1
roadTimer.Interval = roadTimer.Interval - 15
If roadTimer.Interval < 50 Then roadTimer.Interval = 50
riverTimer.Interval = riverTimer.Interval - 15
If riverTimer.Interval < 50 Then riverTimer.Interval = 50
'set amount of traffic
For i = 0 To numOfCars - 1
  imgCar(i).Visible = True
Next i
If level = 1 Then
  imgLog(0).Visible = False
  imgCar(3).Visible = False
  imgCar(6).Visible = False
  imgCar(10).Visible = False
  imgCar(4).Visible = False
  imgCar(0).Visible = False
End If
If level = 2 Then
  imgLog(0).Visible = False
  imgLog(6).Visible = False
  imgCar(3).Visible = False
  imgCar(6).Visible = False
  imgCar(10).Visible = False
End If
If level = 3 Then
  imgLog(0).Visible = False
  imgLog(6).Visible = False
  imgLog(11).Visible = False
  imgCar(3).Visible = False
End If
If level >= 4 Then
  imgLog(0).Visible = False
  imgLog(6).Visible = False
  imgLog(11).Visible = False
End If
For i = 0 To 4
  home(i) = False
  frogState(i) = 0
  imgFrog(i).Visible = False
Next i
retval = PlayWaveFile(App.Path & "\" & "Gribit.wav", True)
newFrog
End Sub


Private Sub scoreTimer_Timer()
lblScore.Caption = Str(score)
If score > highScore Then highScore = score
lblHighScore.Caption = Str(highScore)
lblLevel.Caption = Str(level)
If lives <= 0 Then GoTo 2000
For i = lives - 1 To 4
  imgLife(i).Visible = False
Next i
If lives <= 1 Then GoTo 2000
For i = lives - 2 To 0 Step -1
  imgLife(i).Visible = True
Next i
2000:
End Sub

Private Sub snakeTimer_Timer()
If snakeAlive = False Then
  If Int(Rnd * 10) = 1 Then
    snakeAlive = True
    imgSnake.Left = -1700
    imgSnake.Visible = True
    swap = 0
    swapCount = 0
  End If
Else
  imgSnake.Visible = True
  imgSnake.Left = imgSnake.Left + 80
  If imgSnake.Left > 11000 Then
    imgSnake.Visible = False
    snakeAlive = False
    GoTo 3000
  End If
  swapCount = swapCount + 1
  If swapCount > 1 Then
    swapCount = 0
    If swap = 0 Then
      swap = 1
      imgSnake = LoadPicture(App.Path & "\snake1.gif")
    Else
      swap = 0
      imgSnake = LoadPicture(App.Path & "\snake2.gif")
    End If
  End If
End If
3000:
End Sub
