VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PerifVis 
   AutoRedraw      =   -1  'True
   Caption         =   "Peripheral Vision Game"
   ClientHeight    =   4845
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicSizer 
      AutoSize        =   -1  'True
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdlPerifVis 
      Left            =   720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picBuiltIn 
      Height          =   1380
      Left            =   -120
      Picture         =   "PerifVis.frx":0000
      ScaleHeight     =   1320
      ScaleWidth      =   2160
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.CommandButton cmdSquare 
      Caption         =   "10"
      Height          =   300
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuSubGame 
         Caption         =   "&New"
         Index           =   0
      End
      Begin VB.Menu mnuSubGame 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuSubGame 
         Caption         =   "E&xit"
         Index           =   2
      End
   End
   Begin VB.Menu mnuDiff 
      Caption         =   "&Difficulty"
      Begin VB.Menu mnuLevel 
         Caption         =   "Simple"
         Index           =   0
      End
      Begin VB.Menu mnuLevel 
         Caption         =   "Blank"
         Index           =   1
      End
      Begin VB.Menu mnuLevel 
         Caption         =   "Colour"
         Index           =   2
      End
      Begin VB.Menu mnuLevel 
         Caption         =   "Colour Blank"
         Index           =   3
      End
      Begin VB.Menu mnuLevel 
         Caption         =   "Random colours"
         Index           =   4
      End
      Begin VB.Menu mnuLevel 
         Caption         =   "Blended Colour"
         Index           =   5
      End
      Begin VB.Menu mnuLevel 
         Caption         =   "Picture"
         Index           =   6
         Begin VB.Menu mnuPicLevel 
            Caption         =   "Built-In"
            Index           =   0
         End
         Begin VB.Menu mnuPicLevel 
            Caption         =   "Load Picture"
            Index           =   1
         End
         Begin VB.Menu mnuPicLevel 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuPicLevel 
            Caption         =   "Solve Picture"
            Enabled         =   0   'False
            Index           =   3
         End
      End
      Begin VB.Menu mnuLevel 
         Caption         =   "Button Size"
         Index           =   7
         Begin VB.Menu mnuBSize 
            Caption         =   "Shrink 50%"
            Index           =   0
         End
         Begin VB.Menu mnuBSize 
            Caption         =   "Shrink 10%"
            Index           =   1
         End
         Begin VB.Menu mnuBSize 
            Caption         =   "Grow 10%"
            Index           =   2
         End
         Begin VB.Menu mnuBSize 
            Caption         =   "Grow 200%"
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHlpOpt 
         Caption         =   "How To &Play.."
         Index           =   0
      End
      Begin VB.Menu mnuHlpOpt 
         Caption         =   "Hint"
         Index           =   1
      End
   End
End
Attribute VB_Name = "PerifVis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Copyright 2005 Roger Gichrist
'some code based on aditya8000's 'Visible Light Spectrum' at PSC txtCodeId=61446". Thanks
'inspiration but no code Ryan Spencer's 'A "Button" Form' txtCodeId=61380. Thanks
Private VisCount      As Long
Private HintCount     As Long
Private CurPoints     As Long
Private MaxPoints     As Long
Private DiffLevel     As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub ButtonReSize(ByVal NewWidth As Long)

'set new root button size (note buttons always square)
'then reload the game

  cmdSquare(0).Width = NewWidth
  cmdSquare(0).Height = cmdSquare(0).Width
  Settings True
  NewGame

End Sub

Private Sub CaptionShow()

'display score on from caption

  Me.Caption = "Peripheral Vision Game Score = " & CurPoints & " of " & MaxPoints

End Sub

Private Sub cmdSquare_Click(Index As Integer)

  Dim rndCom As Long
  Dim Hits   As Long
  Dim I      As Long

  If cmdSquare(Index).Tag = "X" Then ' if enabled
    For I = 0 To cmdSquare.Count - 1 'disable all
      cmdSquare(I).Tag = vbNullString
    Next I
    Do
      If VisCount > cmdSquare.Count - Hits Then
'not enough for another round
        ShowScore
        Exit Do
      End If
'select a random button to show
      rndCom = Int(Rnd * cmdSquare.Count)
      If Not cmdSquare(rndCom).Visible Then
        ShowButton rndCom
        Hits = Hits + 1
      End If
    Loop Until Hits = 10
   Else
'deduct a point if the button is not enabled
    CurPoints = CurPoints - 1
  End If
  CaptionShow

End Sub

Private Sub cmdSquare_MouseDown(Index As Integer, _
                                Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

  If Button = vbRightButton Then
    hint
  End If

End Sub

Private Sub Form_Initialize()

  Randomize

End Sub

Private Sub Form_Load()

  On Error Resume Next
  Settings False
  PerifVis.Show
  On Error GoTo 0

End Sub

Private Sub Form_Resize()

  If VisCount > 1 Then
    If MsgBox("Are you sure you want to start a new game?", vbYesNo, "Peripheral Vision Game") = vbYes Then
      NewGame
    End If
  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

  controlArrayUnload cmdSquare
  End

End Sub

Private Sub hint()

' flash the current generation of buttons

  ControlArrayHinting cmdSquare
  CurPoints = CurPoints - 1
  HintCount = HintCount + 1
  CaptionShow
  Sleep 100
  ControlArrayHinting cmdSquare

End Sub

Private Sub LevelBacks()

  Select Case DiffLevel
   Case 0, 1
    Cls
   Case 2, 3, 4, 5
    DrawSpectrum Me, Rnd > 0.5, Rnd > 0.5
'Case 6, 7'Picture stuff do nothing
  End Select

End Sub

Private Sub mnuBSize_Click(Index As Integer)

'Protected from stupid sizes by the code at the
'end of mnuLevel_Click which blocks ridiculously small or large buttons

  Select Case Index
   Case 0
    ButtonReSize cmdSquare(0).Width * 0.5
   Case 1
    ButtonReSize cmdSquare(0).Width * 0.9
   Case 2
    ButtonReSize cmdSquare(0).Width * 1.1
   Case 3
    ButtonReSize cmdSquare(0).Width * 2
  End Select

End Sub

Private Sub mnuHlpOpt_Click(Index As Integer)

'Dim ctrl   As Variant
'Dim OldCap As String

  Select Case Index
   Case 0
    MsgBox "A game to test your peripheral vision." & vbNewLine & _
       "Aim: Fill the form with buttons" & vbNewLine & _
       "1 Click the button to create a new generation of buttons." & vbNewLine & _
       "2 Click any of the newest generation of buttons to create the next genreation" & vbNewLine & _
       "3 Continue until the form is filled." & vbNewLine & _
       "4 Score (displayed on buttons in simpler levels) decreases if you click an older generation of buttons" & vbNewLine & _
       "5 Difficulty: games get more difficult as you go down the menu. You can also resize the buttons to make it easier or harder." & vbNewLine & _
       "6 Difficulty can also be increased by using larger forms. Form resizes to hold picture puzzles." & vbNewLine & _
       "7 Hints (Help menu or Right-Click any button) cost one point, current generation buttons flash." & vbNewLine & _
       "8 Difficulty level and button size are remembered between runs." & vbNewLine & _
       "" & vbNewLine & _
       "Not much of a game but might give you an idea or two." & vbNewLine & _
       "Mostly a doodle suggested by 2 recent PSC uploads," & vbNewLine & _
       " a. Ryan Spencer's 'A ''Button'' Form' txtCodeId=61380" & vbNewLine & _
       " b. aditya8000's 'Visible Light Spectrum' txtCodeId=61446" & vbNewLine & _
       "" & vbNewLine & _
       "Have fun, make improvements", vbInformation, "Peripheral Vision Game"

   Case 1 ' do hinting
    hint
  End Select

End Sub

Private Sub mnuLevel_Click(Index As Integer)

  Dim I As Long

  For I = 0 To mnuLevel.Count - 3
'-1 for 0-based and -2 to avoid the sub-menu headers which can't be Checked
    mnuLevel(I).Checked = I = Index
  Next I
  If Index < mnuLevel.Count - 2 Then '-1 to avoid the  submenus
    Me.Picture = LoadPicture()
    mnuPicLevel(3).Enabled = False
    DiffLevel = Index
    Settings True
    NewGame
  End If
'idot proofing; deactivates menu items if size is not useful
  mnuBSize(0).Enabled = cmdSquare(0).Width * 0.5 > 150
  mnuBSize(1).Enabled = cmdSquare(0).Width * 0.9 > 150
  mnuBSize(2).Enabled = cmdSquare(0).Width * 1.1 <= Me.Width / 5
  mnuBSize(3).Enabled = cmdSquare(0).Width * 2 <= Me.Width / 5

End Sub

Private Sub mnuPicLevel_Click(Index As Integer)

  Select Case Index
   Case 0 ' built-in picture
    DiffLevel = 6
    Settings True
    PicSizer.Picture = picBuiltIn.Picture
    mnuPicLevel(3).Enabled = True
   Case 1 'load a picture fom disk
    DiffLevel = 7
    Settings True
    With cdlPerifVis
      .Filter = "Picture files|*.bmp;*.jpg;*.jpeg|Bitmaps|*.bmp|JPeg|*.jpg;*.jpeg"
      .FilterIndex = 1
      .ShowOpen
      If LenB(.FileName) Then
        Settings True
        PicSizer.Picture = LoadPicture(.FileName)
        mnuPicLevel(3).Enabled = True
      End If
    End With
   Case 3 'Solve the picture(for when you get bored)
    ControlArrayCaptionWipe cmdSquare
    ControlArrayVisible cmdSquare, True
  End Select
  If Index <> 3 Then
    If mnuPicLevel(3).Enabled Then
      Me.Picture = PicSizer.Picture
      Me.Move Me.Left, Me.Top, PicSizer.Width, PicSizer.Height
      NewGame
    End If
  End If

End Sub

Private Sub mnuSubGame_Click(Index As Integer)

  Select Case Index
   Case 0 '&New
    NewGame
   Case 2 'E&xit
    Unload Me
  End Select

End Sub

Private Sub NewGame()

'generate a new array of buttons

  Dim J          As Integer
  Dim CN         As Integer
  Dim I          As Long
  Dim WF         As Long
  Dim HF         As Long
  Static Working As Boolean   'stops multiple hits on this routine

  Me.Caption = "Peripheral Vision Game"
  If Not Working Then
    Working = True
    LevelBacks
'clear out the old array
    controlArrayUnload cmdSquare
'get a roughly size to the form to the button
    WF = PerifVis.ScaleWidth / cmdSquare(0).Width
    HF = PerifVis.ScaleHeight / cmdSquare(0).Height
    cmdSquare(0).Visible = False
'rescale button to fit exactly on the form
    cmdSquare(0).Move 0, 0, PerifVis.ScaleWidth / WF, PerifVis.ScaleHeight / HF
    Settings True
    ControlArrayGridLoad cmdSquare, WF, HF
'colour buttons depending on Difficulty level
    For I = 0 To WF
      DoEvents
      For J = 0 To HF
        Select Case DiffLevel
         Case 0, 1 'basic button face
          cmdSquare(CN).BackColor = vbButtonFace
         Case 2, 3 ' ordered rows of spectrum
          cmdSquare(CN).BackColor = SpectralColour(400 + (I * 100) Mod 300)
         Case 4 ' random spectrum colours
          cmdSquare(CN).BackColor = SpectralColour((Rnd * 300) + 400)
         Case 5
 'get colour of form at centre of button (engage the Spectrum code; could be replaced by Point, here for fun)
          cmdSquare(CN).BackColor = SpectralColour(SpectralValue(Me.Point(cmdSquare(CN).Left + cmdSquare(CN).Height / 2, cmdSquare(CN).Top + cmdSquare(CN).Width / 2)))
         Case 6, 7
 'Picture puzzle modes gets colour of form at centre of button(standard Point method)
          cmdSquare(CN).BackColor = Me.Point(cmdSquare(CN).Left + cmdSquare(CN).Height / 2, cmdSquare(CN).Top + cmdSquare(CN).Width / 2)
        End Select
        CN = CN + 1
      Next J
    Next I
    Working = False
    VisCount = 0
    HintCount = 0
    CurPoints = cmdSquare.Count \ 10
    MaxPoints = CurPoints
    DoEvents
    ShowButton Int(Rnd * cmdSquare.Count) ' turn on one random button
  End If
  On Error GoTo 0

End Sub

Private Function ScoreComment(Optional ByVal bLegit As Boolean = True) As String

'final score comments

  If bLegit Then
    Select Case CurPoints
     Case MaxPoints
      If HintCount = 0 Then
        ScoreComment = "Perfect! " & vbNewLine & "For a harder game increase the form size."
       Else
        ScoreComment = "Well played." & vbNewLine & "For a harder game increase the form size."
      End If
     Case Is < 0
      ScoreComment = "Try not to focus too hard on the screen!"
     Case Is < MaxPoints * 0.75
      ScoreComment = "Could do better."
     Case Is < MaxPoints * 0.5
      ScoreComment = "Not great."
     Case Is < MaxPoints * 0.25
      ScoreComment = "Very weak."
     Case Else
      ScoreComment = "Not bad try again."
    End Select
    ScoreComment = " Score = " & CurPoints & " of " & MaxPoints & IIf(HintCount, "  Hints: " & HintCount, vbNullString) & vbNewLine & _
                   ScoreComment & vbNewLine & _
                   vbNewLine & _
                   "Play again?"
   Else
    ScoreComment = "Start a New Game?"
  End If

End Function

Private Sub Settings(ByVal bSaveTLoadF As Boolean)

  Dim strFName As String

  If bSaveTLoadF Then
    SaveSetting "PerifVis", "Settings", "Diff", DiffLevel
    If LenB(cdlPerifVis.FileName) Then
      SaveSetting "PerifVis", "Settings", "PrevPic", cdlPerifVis.FileName
    End If
    SaveSetting "PerifVis", "Settings", "Bsize", cmdSquare(0).Width
   Else
    DiffLevel = GetSetting("PerifVis", "Settings", "Diff", 0)
    strFName = GetSetting("PerifVis", "Settings", "PrevPic", vbNullString)
' must be last laoded becuase it recurses to this Sub in Save Mode and would wipe other settings
    ButtonReSize GetSetting("PerifVis", "Settings", "BSize", 300)
    If LenB(Dir(strFName)) Then
      cdlPerifVis.FileName = strFName
    End If
    Select Case DiffLevel
     Case Is < 6
      mnuLevel_Click CInt(DiffLevel)
     Case 6
      mnuPicLevel_Click 0
     Case 7
      If LenB(Dir(strFName)) Then
        PicSizer.Picture = LoadPicture(strFName)
        mnuPicLevel(3).Enabled = True
        Me.Picture = PicSizer.Picture
        Me.Move Me.Left, Me.Top, PicSizer.Width, PicSizer.Height
        NewGame
       Else 'picture is gone so open cCmmenDialog
        mnuPicLevel_Click 1
      End If
    End Select
  End If

End Sub

Private Sub ShowButton(ByVal ID As Long)

  VisCount = VisCount + 1 ' count visible buttons
  With cmdSquare(ID)
    .Visible = True                         'make button visible
    .Tag = "X"                              'enable button for current generation of buttons
    If DiffLevel = 0 Or DiffLevel = 2 Then  'caption it (or not)
      .Caption = CurPoints
     Else
      .Caption = vbNullString
    End If
  End With

End Sub

Private Sub ShowScore(Optional bLegit As Boolean = True)

  Dim I      As Long

  If bLegit Then
    For I = 0 To cmdSquare.Count - 1
'show remaining buttons if the leftovers are not enough for a new set
      If Not cmdSquare(I).Visible Then
        ShowButton I
      End If
    Next I
  End If
  If MsgBox(ScoreComment(bLegit), vbYesNo, "Peripheral Vision Game") = vbYes Then
    NewGame
   Else
    Unload Me
  End If

End Sub


':)Code Fixer V4.0.0 (Tuesday, 05 July 2005 11:41:09) 10 + 382 = 392 Lines Thanks Ulli for inspiration and lots of code.
':)SETTINGS DUMP: 13330232222333323|333320222222222222222222222222|1112222|2221222|222222222233|1111111111111|1122222222222|333333|

