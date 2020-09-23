VERSION 5.00
Begin VB.Form GForm 
   BorderStyle     =   0  'None
   Caption         =   "VB DOOM - by Simon Price - www.VBgames.co.uk"
   ClientHeight    =   5772
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7692
   ForeColor       =   &H0000FF00&
   Icon            =   "GForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "GForm.frx":030A
   ScaleHeight     =   481
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   641
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame CollisionF 
      Caption         =   "Collision Detection"
      Height          =   972
      Left            =   1440
      TabIndex        =   27
      Top             =   2400
      Visible         =   0   'False
      Width           =   4812
      Begin VB.CheckBox CollisionO 
         Caption         =   "Collision Detection"
         Height          =   192
         Left            =   1560
         TabIndex        =   29
         Top             =   720
         Width           =   1692
      End
      Begin VB.Label Label3 
         Caption         =   "The collision detection at the moment is a bit dodgy, so you can turn this off if you want, allowing you to move more easily."
         Height          =   372
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   4572
      End
   End
   Begin VB.Frame LevelF 
      Caption         =   "Level Selection"
      Height          =   2172
      Left            =   1440
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   4812
      Begin VB.OptionButton LevelO 
         Caption         =   "Level 10 (hardest)"
         Height          =   252
         Index           =   10
         Left            =   2640
         TabIndex        =   26
         Top             =   1800
         Width           =   1692
      End
      Begin VB.OptionButton LevelO 
         Caption         =   "Level 9"
         Height          =   252
         Index           =   9
         Left            =   2640
         TabIndex        =   25
         Top             =   1560
         Width           =   1692
      End
      Begin VB.OptionButton LevelO 
         Caption         =   "Level 8"
         Height          =   252
         Index           =   8
         Left            =   2640
         TabIndex        =   24
         Top             =   1320
         Width           =   1692
      End
      Begin VB.OptionButton LevelO 
         Caption         =   "Level 7"
         Height          =   252
         Index           =   7
         Left            =   2640
         TabIndex        =   23
         Top             =   1080
         Width           =   1692
      End
      Begin VB.OptionButton LevelO 
         Caption         =   "Level 6"
         Height          =   252
         Index           =   6
         Left            =   2640
         TabIndex        =   22
         Top             =   840
         Width           =   1692
      End
      Begin VB.OptionButton LevelO 
         Caption         =   "Level 5"
         Height          =   252
         Index           =   5
         Left            =   240
         TabIndex        =   21
         Top             =   1800
         Width           =   1692
      End
      Begin VB.OptionButton LevelO 
         Caption         =   "Level 4"
         Height          =   252
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Top             =   1560
         Width           =   1692
      End
      Begin VB.OptionButton LevelO 
         Caption         =   "Level 3"
         Height          =   252
         Index           =   3
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Width           =   1692
      End
      Begin VB.OptionButton LevelO 
         Caption         =   "Level 2"
         Height          =   252
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   1692
      End
      Begin VB.OptionButton LevelO 
         Caption         =   "Level 1 (easiest)"
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Value           =   -1  'True
         Width           =   1692
      End
      Begin VB.Label Label2 
         Caption         =   $"GForm.frx":32FEE
         Height          =   612
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   4572
      End
   End
   Begin VB.Frame StyleF 
      Caption         =   "Graphics Style"
      Height          =   1572
      Left            =   1440
      TabIndex        =   11
      Top             =   3480
      Visible         =   0   'False
      Width           =   4812
      Begin VB.OptionButton RealO 
         Caption         =   "Realistic - Detailed textures - brick, stone, wood, leather etc."
         Height          =   192
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Value           =   -1  'True
         Width           =   4452
      End
      Begin VB.OptionButton WierdO 
         Caption         =   "Wierd stuff - brightly colored, patterned walls"
         Height          =   192
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   3492
      End
      Begin VB.Label Label1 
         Caption         =   "Please select your preferred style of graphics:"
         Height          =   252
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   4332
      End
   End
   Begin VB.PictureBox WallPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1212
      Index           =   8
      Left            =   4560
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   501
      TabIndex        =   10
      Top             =   4680
      Visible         =   0   'False
      Width           =   6012
   End
   Begin VB.PictureBox WallPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1212
      Index           =   7
      Left            =   2880
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   501
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   6012
   End
   Begin VB.PictureBox WallPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1212
      Index           =   6
      Left            =   5640
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   501
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   6012
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   492
      Left            =   3240
      TabIndex        =   7
      Top             =   5160
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.PictureBox WallPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1212
      Index           =   5
      Left            =   5880
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   501
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   6012
   End
   Begin VB.PictureBox WallPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1212
      Index           =   4
      Left            =   6120
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   501
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   6012
   End
   Begin VB.PictureBox WallPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1212
      Index           =   3
      Left            =   6360
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   501
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   6012
   End
   Begin VB.PictureBox WallPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1212
      Index           =   2
      Left            =   6600
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   501
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   6012
   End
   Begin VB.PictureBox WallPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1212
      Index           =   1
      Left            =   3600
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   501
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   6012
   End
   Begin VB.PictureBox LevelPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   240
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   6360
      Picture         =   "GForm.frx":33079
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   4800
   End
End
Attribute VB_Name = "GForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Level As tLevel 'the level
Dim FrameCount As Integer 'I used this to test the fps
Dim StopFlashing As Boolean 'true tells the start message to stop flashing
Dim DoorPos As tCoOrd 'exit co-ords
Dim LevelNo As Byte 'level chosen
Dim Collisions As Boolean 'if collision detection is on or not

Private Sub cmdStart_Click()
Dim Path As String
'check wot graphics to load up
If RealO.Value Then
  Path = App.Path & "\Real\"
Else
  Path = App.Path & "\Wierd\"
End If
'load the chosen pics
For i = 1 To 8
  WallPic(i) = LoadPicture(Path & i & ".bmp", , &H2)
Next
'load level map
LevelPic = LoadPicture(App.Path & "\Levels\" & LevelNo & ".bmp")
'make options dissapear
LevelF.Visible = False
CollisionF.Visible = False
StyleF.Visible = False
cmdStart.Visible = False
'set collsion detection
If CollisionO.Value Then Collisions = True
'load the the level
LoadLevel
'enter main loop
MainLoop
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyS 'go to options screen
    DoOptions
  Case vbKeyEscape
  Hide 'exit the game
  MsgBox "Thankyou for playing this early version of VB Doom by Simon Price. To see more games (some by me) visit my website : www.VBgames.co.uk - send any feedback on this game to : gamefeedback@VBgames.co.uk", vbInformation, "Thanks 4 Playing!"
  MsgBox "The game will now attempt to revert back to your original screen resolution, however, it is not always successful so you may have to do this manually.", vbInformation, "Reverting to original screen res"
Select Case iWidth 'put screen res back to normal
Case 800
ChangeScreenSettings 800, 600, 16
Case 1024
ChangeScreenSettings 1024, 768, 16
End Select
    End
  Case vbKeyC 'capture the screen
    PB.Picture = PB.Image
    SavePicture PB.Picture, App.Path & "\Pic.bmp"
End Select
End Sub

Public Sub DoOptions()
StopFlashing = True
'clear title screen
Picture = LoadPicture()
Cls
'show all the options
LevelF.Visible = True
CollisionF.Visible = True
StyleF.Visible = True
cmdStart.Visible = True
End Sub

Private Sub Form_Load()
Randomize Timer
'remember the screen res b4 messing with it
RememberScreenRes
'change to low-res, 16 bit color
ChangeScreenSettings 640, 480, 16

Show
'build TONS of look-up tables
RememberStuff
'sort out the forms layout
Move 0, 0, RAYSby2 * Screen.TwipsPerPixelX, RAYSby1andHALF * Screen.TwipsPerPixelY
PB.Move 0, 0, RAYSby2, RAYSby1andHALF
LevelNo = 1
'keep flashing a message
Do
For i = 1 To 10000
DoEvents
Next
Print " "
Print "Press 'S' to begin"
For i = 1 To 10000
DoEvents
Next
Cls
Loop Until StopFlashing
End Sub

Public Sub MainLoop()
On Error Resume Next

Dim x As Long
Dim y As Long
Dim Temp As POINTAPI
Dim RayAngle As Single
Dim ScrX As Integer
Dim StepX As Integer
Dim StepY As Integer
Dim Length As Integer
Dim Hit As Byte
Dim DarkX As Integer

LoadLevel

Do
DoEvents

FrameCount = FrameCount + 1
Caption = FrameCount

PB.Cls
'walk forward
If (GetKeyState(vbKeyUp) And KEY_DOWN) Then
      Man.x = Man.x + Cosine(Man.Angle + ADD90DEGREES) / 10
      Man.y = Man.y - Sine(Man.Angle + ADD90DEGREES) / 10
      If Level.Tile(Man.x, Man.y) = 8 Then Exit Do
      If Collisions Then 'check for walls
      If Level.Tile(Man.x, Man.y) <> NOWT Then
        Man.x = Man.x - Cosine(Man.Angle + ADD90DEGREES) / 10
        Man.y = Man.y + Sine(Man.Angle + ADD90DEGREES) / 10
      End If
      End If
End If
'walk backwards
If (GetKeyState(vbKeyDown) And KEY_DOWN) Then
      Man.x = Man.x - Cosine(Man.Angle + ADD90DEGREES) / 10
      Man.y = Man.y + Sine(Man.Angle + ADD90DEGREES) / 10
      If Level.Tile(Man.x, Man.y - 0.5) = 8 Then Exit Do
      If Collisions Then 'check for walls
      If Level.Tile(Man.x, Man.y - 0.5) <> NOWT Then
        Man.x = Man.x + Cosine(Man.Angle + ADD90DEGREES) / 10
        Man.y = Man.y - Sine(Man.Angle + ADD90DEGREES) / 10
      End If
      End If
End If
'turn left
If (GetKeyState(vbKeyLeft) And KEY_DOWN) Then
    If Man.Angle < 0 Then
      Man.Angle = BACKHALFVIEW
    Else
      Man.Angle = Man.Angle - TURNANGLE
    End If
End If
'turn right
If (GetKeyState(vbKeyRight) And KEY_DOWN) Then
    If Man.Angle > BACKHALFVIEW Then
      Man.Angle = 0
    Else
      Man.Angle = Man.Angle + TURNANGLE
    End If
End If
'this set the first ray 30 degrees to the left of view
RayAngle = Man.Angle - HALFVIEWRAYS

'loop through all 320 rays, drawing a slither of screen each time
For ScrX = 0 To RAYS
  
  x = Man.x * 1200000 'multiply up so that the fixed-point maths is
  y = Man.y * 1200000 'accurate enough
  StepX = Sine(RayAngle) / RAYDETAIL * 1200000 'i.e. 1/10th of a unit
  StepY = Cosine(RayAngle) / RAYDETAIL * 1200000
  Length = 0 'length of ray is reset
  
  Do
    x = x - StepX
    y = y - StepY 'move ray along a bit
    Length = Length + 1 'increment length
    Hit = Level.Tile(x \ 1200000, y \ 1200000) 'see wot's hit
  Loop Until Hit 'keep the ray going until a hit is detected
  
  DarkX = Dist2Dark(Length) 'see how dark the wall should be
  Length = Dist2Height(Length) 'and how tall it looks based on ray length
      
  Temp.x = (x Mod 1200000) \ 12000
  Temp.y = (y Mod 1200000) \ 12000 'scale stuff back down again
  
  'here's the clever bit that no-ones ever done b4
  'perspective textures are put onto a wall which
  'was only represented by 1 byte of memory - now that's
  'efficient!!! The drawback is that it's only 90% accurate
  '-this technique gives incorrect results at the sides of walls
  'but for textured 3d walls in VB I think we can forget that
  'since it's hardly noticable anyway
  If Abs(50 - Temp.x) < Abs(50 - Temp.y) Then
    StretchBlt PB.hdc, ScrX, RAYSby3div8 - Length, 1, Length + Length, WallPic(Hit).hdc, Temp.x + DarkX, 0, 1, 100, vbSrcCopy
  Else
    StretchBlt PB.hdc, ScrX, RAYSby3div8 - Length, 1, Length + Length, WallPic(Hit).hdc, Temp.y + DarkX, 0, 1, 100, vbSrcCopy
  End If
'fire next ray 1 pixel further along
RayAngle = RayAngle + 1
Next
'copy from backbuffer
StretchBlt hdc, 0, 0, RAYSby2, RAYSby1andHALF, PB.hdc, 0, 0, RAYS, RAYSby3div4, vbSrcCopy

Loop

'level complete
If MsgBox("Well done, you have completed level " & LevelNo & ". Do want to play again? Click 'Yes' to pick another level and play again, or click 'No' to exit the game.", vbQuestion + vbYesNo, "Level Complete") = vbYes Then
  Cls
  DoOptions
Else
  MsgBox "Thankyou for playing aMAZEing 3D by Simon Price. See more cool VB games at www.VBgames.co.uk", vbInformation, "Thanks 4 Playing!"
  Unload Me
End If

End Sub

Public Sub LoadLevel()
Dim x As Byte
Dim y As Byte

Man.Angle = 0
'loads a level by transferring bitmap into memory
Level.Size = LevelPic.Width - 1
ReDim Level.Tile(0 To Level.Size, 0 To Level.Size)

For x = 0 To Level.Size
For y = 0 To Level.Size
  Select Case GetPixel(GForm.LevelPic.hdc, x, y)
    Case vbBlack
      Level.Tile(x, y) = NOWT
    Case vbCyan
      Level.Tile(x, y) = 1
    Case vbYellow
      Level.Tile(x, y) = 2
    Case vbBlue
      Level.Tile(x, y) = 3
    Case QBColor(6)
      Level.Tile(x, y) = 4
    Case QBColor(7)
      Level.Tile(x, y) = 5
    Case vbMagenta
      Level.Tile(x, y) = 6
    Case vbWhite
      Level.Tile(x, y) = 7
    Case vbGreen
      Level.Tile(x, y) = NOWT
      Man.x = x
      Man.y = y
    Case vbRed
      Level.Tile(x, y) = 8
      DoorPos.x = x
      DoorPos.y = y
  End Select
Next

Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
'change screen res back 2 norm
Select Case iWidth
Case 800
ChangeScreenSettings 800, 600, 16
Case 1024
ChangeScreenSettings 1024, 768, 16
End Select
End Sub

Private Sub LevelO_Click(Index As Integer)
LevelNo = Index
End Sub
