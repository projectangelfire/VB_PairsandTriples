VERSION 5.00
Begin VB.Form Pairs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAIRS!"
   ClientHeight    =   8250
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11220
   DrawStyle       =   5  'Transparent
   FillColor       =   &H00008000&
   FillStyle       =   5  'Downward Diagonal
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Pairs.frx":0000
   ScaleHeight     =   8250
   ScaleMode       =   0  'User
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2760
      Top             =   2880
   End
   Begin VB.CommandButton btnCard 
      BackColor       =   &H00E0E0E0&
      Height          =   1695
      Index           =   0
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblTimeLimit 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   7560
      Width           =   4575
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   4440
      TabIndex        =   1
      Top             =   7440
      Width           =   2775
   End
   Begin VB.Label lblInstruction 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   7320
      TabIndex        =   0
      Top             =   7560
      Width           =   3975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Index           =   1
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New Game"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close Game"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuGameOptions 
         Caption         =   "Game Options"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHowTo 
         Caption         =   "&How to Play"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Game"
      End
   End
End
Attribute VB_Name = "Pairs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ********************************************************************************
' *                                                                              *
' *                            Pairs & Triples Game                              *
' *            Designed by AB for FdSc Applied Computing                  *
' *                             Copyright(c), 2006                               *
' *                                Final Build                                   *
' ********************************************************************************
Option Explicit

' Load a library that lets VB recognise 'long file names' used by Win XP
' Kernel32.dll - Core Windows 32-bit API support
Private Declare Function GetShortPathName Lib "kernel32" Alias _
"GetShortPathNameA" (ByVal lpszLongPath As String, _
ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Sub Sleep Lib "kernel32" _
 (ByVal dwMilliseconds As Long) ' Also load from Kernel32, Sleep function

' ************************ Declare Global Variables *****************************
Dim blnSelectedA() As Boolean, blnSelectedB() As Boolean, _
blnSelectedC() As Boolean ' Track Selected Cards
Dim blnTotalSelected() As Boolean
Public intPlayerAScore As Integer ' These are self-explanatory
Public intPlayerBScore As Integer
Public blnPlayerTurn As Boolean   ' A = True B = False
Public intCardCounter As Integer ' Number of Cards selected that turn
Public intTotalNumber As Integer ' The total number of cards selected by the user
Public blnPairs As Boolean ' True = Pairs, False = Trebles
Public intNoPairs As Integer 'Number of distinct cards
Public intTimeLimit As Integer 'Time Limit per turn, 0 = none
Public intTimePassed As Integer ' Keep track of how much time a player has had
Public intTotalSelect As Integer ' Tallys the total number of cards selected
Public strCardType As String ' Gets the cards that need to be loaded
' *******************************************************************************

Private Sub InitialiseGame()

Call DetermineGameSettings ' Gets the settings from the options file
Call LoadCards 'Loads the cards into memory based on those settings
Call ResetVariables 'Resets all the game variables to null
Pairs.Refresh ' Refreshes the form just to get rid of anything unwanted
intTotalSelect = 0 ' Resets the total number of pairs obtained by both players

' This routine shuffles the cards and places them in a random order
Dim intCarPos() As Integer
ReDim intCarPos(intTotalNumber) 'Dumb ass Microsoft
Dim intCounter As Integer
Dim intChecker As Integer
Randomize Timer
For intCounter = 1 To intTotalNumber ' This routine juggles the cards
redo:
    intCarPos(intCounter) = Int(Rnd * intTotalNumber) + 1
        For intChecker = 1 To intCounter
            If intCounter = intChecker Then Exit For
            If intCarPos(intCounter) = intCarPos(intChecker) Then
                GoTo redo ' bad evil code here
            End If
        Next intChecker
        ' Print intCounter & " " & intCarPos(intCounter) ' for Design time debugging
btnCard(intCarPos(intCounter)).TabIndex = intCounter ' Prevents people from cheating
Next intCounter

' Card Shuffling Routine
' This works out the positions of the cards on screen and displays them
Dim intWidth As Integer, intHeight As Integer
intHeight = DetermineGrid() ' Calls a routine that decides the best card layout
intWidth = RoundUp(intTotalNumber / intHeight) ' should return an integer

Dim intNWid As Integer ' This is for the unique spacing of the 3rd & 4th line
Dim intFKH As Integer
    intFKH = 7000 ' Fake Height of the Form
For intCounter = 1 To intTotalNumber
    Select Case intCounter
        Case Is <= intWidth
            With btnCard(intCarPos(intCounter))
                .Left = ((Me.Width / intWidth) * (intCounter - 1)) + _
                    ((Me.Width / intWidth) / intWidth)
                .Top = (intFKH / intHeight) * 0 + ((intFKH / intHeight) / intHeight)
            End With
        Case intWidth To (intWidth * 2)
            With btnCard(intCarPos(intCounter))
                .Left = ((Me.Width / intWidth) * ((intCounter - intWidth) - 1)) + _
                    ((Me.Width / intWidth) / intWidth)
                .Top = (intFKH / intHeight) * 1 + ((intFKH / intHeight) / intHeight)
            End With
        Case (intWidth * 2) To (intWidth * 3)
                If intTotalNumber < intWidth * 3 Then
                    intNWid = intTotalNumber - (intWidth * 2)
                        Else
                    intNWid = intWidth
                End If
            With btnCard(intCarPos(intCounter))
                .Left = ((Me.Width / intNWid) * ((intCounter - (intWidth * 2)) - 1)) + _
                    ((Me.Width / intNWid) / intNWid)
                .Top = (intFKH / intHeight) * 2 + ((intFKH / intHeight) / intHeight)
            End With
        Case Is > (intWidth * 3) ' Special Case
            intNWid = intTotalNumber - (intWidth * 3)
            With btnCard(intCarPos(intCounter))
                .Left = ((Me.Width / intNWid) * ((intCounter - (intWidth * 3)) - 1)) + _
                    ((Me.Width / intNWid) / intNWid)
                .Top = (intFKH / intHeight) * 3 + ((intFKH / intHeight) / intHeight)
            End With
        End Select
Next intCounter

' Make the Cards Visisble
For intCounter = 1 To intTotalNumber: btnCard(intCounter).Visible = True: Next intCounter
' For intCounter = 1 To intTotalNumber: btnCard(intCounter).Caption = CStr(intCounter): Next intCounter
Call UpdateBoxes
End Sub

Private Sub btnCard_Click(Index As Integer)
tmrDelay.Enabled = True
If blnTotalSelected(Index) = True Then Exit Sub 'If already clicked then ignore
blnTotalSelected(Index) = True ' Remembers that this image has already been clicked
intCardCounter = intCardCounter + 1 'Adds a notch to the selected number of cards
'Print CardSet(Index) & "-" & DetermineIndex(Index) & "---" & intCardCounter
Call AddToSelected(Index) ' Remembers the card has been chosen
Call LoadImage(Index, DetermineIndex(Index)) ' shows this on the screen with new image
If CheckPairs(DetermineIndex(Index)) = True Then ' Determines if the player has matching cards
    'Call DisplayText("PAIR!!", Pairs, 200, 200, &H8846&, "Arial", 72)
    intTimePassed = 0 ' Resets the Timer
    Call UpdateGame(DetermineIndex(Index)) 'Gets rid of the pairs
    Call UpdateScore 'Adds a point to players score
    Call NextTurn ' Resets the game for the next turn
    Call UpdateBoxes 'Updates scores etc
    If intTotalSelect = intTotalNumber Then Call EndGame
Else
Call GameType(False) ' No Pairs have been selected so the game resets for the next player
End If
End Sub

Private Sub AddToSelected(Index As Integer) ' Adds selected cards to memory
Select Case blnPairs
    Case Is = True
        If Index <= intTotalNumber / 2 Then
            blnSelectedA(Index) = True
                ElseIf Index > intTotalNumber / 2 Then
            blnSelectedB(Index - (intTotalNumber / 2)) = True
        End If
    Case Is = False
        If Index <= intTotalNumber / 3 Then
            blnSelectedA(Index) = True
                ElseIf Index > intTotalNumber / 3 And _
                    Index <= ((intTotalNumber / 3) * 2) Then
                        blnSelectedB(Index - (intTotalNumber / 3)) = True
                ElseIf Index > (intTotalNumber / 3) * 2 Then
                        blnSelectedC(Index - ((intTotalNumber / 3) * 2)) = True
        End If
End Select
End Sub

Private Function DetermineIndex(Index As Integer) As Integer
Select Case blnPairs
    Case Is = True
          If Index <= intTotalNumber / 2 Then
                DetermineIndex = Index
                ElseIf Index > intTotalNumber / 2 Then
            DetermineIndex = (Index - (intTotalNumber / 2))
        End If
    Case Is = False
        If Index <= intTotalNumber / 3 Then
            DetermineIndex = Index
                ElseIf Index > intTotalNumber / 3 And _
                    Index <= ((intTotalNumber / 3) * 2) Then
                        DetermineIndex = (Index - (intTotalNumber / 3))
                    ElseIf Index > (intTotalNumber / 3) * 2 Then
                        DetermineIndex = (Index - ((intTotalNumber / 3) * 2))
        End If
End Select
End Function

Private Function CheckPairs(Index As Integer) As Boolean
Select Case blnPairs
    Case Is = True
        If blnSelectedA(Index) = True And blnSelectedB(Index) = True Then
            CheckPairs = True
            intTotalSelect = intTotalSelect + 2
        End If
    Case Is = False
        If blnSelectedA(Index) = True And blnSelectedB(Index) = True _
            And blnSelectedC(Index) = True Then
                CheckPairs = True
                intTotalSelect = intTotalSelect + 3
        End If
End Select
End Function

Private Function CardSet(Index As Integer) As Integer
Select Case blnPairs
    Case Is = True
        If Index <= intTotalNumber / 2 Then
            CardSet = 1
        ElseIf Index > intTotalNumber / 2 Then
            CardSet = 2
        End If
    Case Is = False
        If Index <= intTotalNumber / 3 Then
            CardSet = 1
                ElseIf Index > intTotalNumber / 3 And _
                    Index <= ((intTotalNumber / 3) * 2) Then
                        CardSet = 2
                    ElseIf Index > (intTotalNumber / 3) * 2 Then
                        CardSet = 3
        End If
End Select
End Function

Private Function GetShortName(ByVal sLongFileName As String) As String
Dim lRetVal As Long, sShortPathName As String, iLen As Integer
'Buffer area for API function call return
sShortPathName = Space(255)
iLen = Len(sShortPathName)
lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
GetShortName = Left(sShortPathName, lRetVal)
End Function

Private Sub NextTurn()
Dim intCounter As Integer
    Select Case blnPairs
        Case Is = True
            For intCounter = 1 To intTotalNumber / 2
                blnSelectedA(intCounter) = False
                blnSelectedB(intCounter) = False
            Next intCounter
        Case Is = False
            For intCounter = 1 To intTotalNumber / 3
                blnSelectedA(intCounter) = False
                blnSelectedB(intCounter) = False
                blnSelectedC(intCounter) = False
            Next intCounter
    End Select

If CheckExist(GetShortName(CurDir() & "\graphics\" & strCardType & "\" & _
    strCardType & ".jpg")) = True Then
    For intCounter = 1 To intTotalNumber
        btnCard(intCounter).Picture = LoadPicture(GetShortName(CurDir() & _
            "\graphics\" & strCardType & "\" & strCardType & ".jpg"))
        blnTotalSelected(intCounter) = False
    Next intCounter
End If
Wait 1.3
intCardCounter = 0 ' Resets the selected cards to 0
tmrDelay.Enabled = False
' The cards need refreshing otherwise a glitch CAN occur, if someone is trying to
' MAKE IT SO
For intCounter = 1 To intTotalNumber
    btnCard(intCounter).Refresh
Next intCounter
End Sub

Private Sub UpdateGame(intRealIndex)
Select Case blnPairs
    Case Is = True
        btnCard(intRealIndex).Visible = False
        btnCard(intRealIndex + (intTotalNumber / 2)).Visible = False
    Case Is = False
        btnCard(intRealIndex).Visible = False
        btnCard(intRealIndex + (intTotalNumber / 3)).Visible = False
        btnCard(intRealIndex + ((intTotalNumber / 3) * 2)).Visible = False
End Select
End Sub

Private Sub EndGame()
lblInstruction.Caption = "Game Over"
lblInstruction.Refresh
If intPlayerAScore > intPlayerBScore Then
    lblTimeLimit.Caption = "Player A Wins!"
        ElseIf intPlayerAScore < intPlayerBScore Then
    lblTimeLimit.Caption = "Player B Wins!"
        ElseIf intPlayerAScore = intPlayerBScore Then
    lblTimeLimit.Caption = "The Game is a Draw!"
End If
End Sub

Private Sub DetermineGameSettings()
If CheckExist(GetShortName(CurDir() & "\options.dat")) = True Then
Dim intFree As Integer
    intFree = FreeFile()
Open GetShortName(CurDir() & "\options.dat") For Random As #intFree
    Get #intFree, 1, intNoPairs
    Get #intFree, 2, blnPairs
    Get #intFree, 3, intTimeLimit
    Get #intFree, 5, strCardType
Close #1
Else: End
End If

Select Case blnPairs
    Case Is = True
        intTotalNumber = intNoPairs * 2
    Case Is = False
        intTotalNumber = intNoPairs * 3
End Select
End Sub

Private Sub LoadCards() ' Loads the cards to be used onto the form
On Error Resume Next 'Bad bad way of solving this problem
Dim intCounter As Integer
If CheckExist(GetShortName(CurDir() & "\graphics\" & strCardType & _
    "\" & strCardType & ".jpg")) = True Then
For intCounter = 1 To intTotalNumber
    Load btnCard(intCounter)
        With btnCard(intCounter)
            .Picture = LoadPicture(GetShortName(CurDir() & "\Graphics\" & strCardType & _
                "\" & strCardType & ".jpg"))
            .Height = 1695 'This is the only place to adjust the height and width of the cards
            .Width = 1400 'formerly 1095
        End With
Next intCounter
End If
End Sub

Private Sub UnloadCards() 'Gets rid of any cards still in memory
Dim intCounter As Integer
If intTotalNumber = 0 Then Exit Sub
For intCounter = 1 To intTotalNumber
    Unload btnCard(intCounter)
Next intCounter
End Sub

Private Sub LoadImage(intIndex As Integer, intImageNo As Integer)
Dim strDirectory As String
strDirectory = GetShortName(CurDir())
If CheckExist(GetShortName(strDirectory & "\graphics\" & strCardType & _
    "\" & CStr(intImageNo) & ".jpg")) = True Then
Dim intno As Integer
intno = CStr(intIndex)
btnCard(intIndex).Picture = LoadPicture(strDirectory & "\Graphics\" & _
    strCardType & "\" & CStr(intImageNo) & ".jpg")
btnCard(intIndex).Refresh ' Bit annoying trying to find this statement
End If
End Sub

Private Sub Wait(Seconds As Single) 'Nice little sub for program pauses
    Dim lngMilliSeconds As Long
    lngMilliSeconds = Seconds * 1000
    Sleep lngMilliSeconds
End Sub

Private Sub Form_Load()
If CheckExist(GetShortName((CurDir() & "\options.dat"))) = True Then
    Dim strPicture As String
    Dim intFree As Integer
        intFree = FreeFile()
Open GetShortName(CurDir() & "\options.dat") For Random As #intFree
       Get #intFree, 4, strPicture
Close #intFree
End If
If CheckExist(GetShortName(CurDir() & "\graphics\BackDrop\" & _
    strPicture & ".jpg")) = True Then
Pairs.Picture = LoadPicture(GetShortName(CurDir() & "\graphics\BackDrop\" & _
    strPicture & ".jpg"))
End If
End Sub

Private Sub tmrDelay_Timer()
If intTimeLimit = 0 Then Exit Sub
intTimePassed = intTimePassed + 1
lblTimeLimit.Caption = "Time Remaining: " & intTimeLimit - intTimePassed
    If intTimePassed >= intTimeLimit Then
        lblTimeLimit.Caption = "Time has Expired!"
        lblTimeLimit.Refresh
        intTimePassed = 0
        Call GameType(True)
    End If
End Sub

Private Function CheckExist(strFile As String) As Boolean
If Dir$(strFile) <> "" Then
    CheckExist = True
Else
    CheckExist = False
    MsgBox ("A file required by pairs appears to have become corrupted " & vbCrLf & _
        "or missing. Please reinstall the application.")
    End
End If
End Function

Private Sub ResetVariables()
' Resets all the varibales of the game to their defaults
' Dimension the selected card variables accoring to gametype
ReDim blnTotalSelected(intTotalNumber)
If blnPairs = True Then
    ReDim blnSelectedA(intTotalNumber / 2)
    ReDim blnSelectedB(intTotalNumber / 2)
        Else
    ReDim blnSelectedA(intTotalNumber / 3)
    ReDim blnSelectedB(intTotalNumber / 3)
    ReDim blnSelectedC(intTotalNumber / 3)
End If

intPlayerAScore = 0: intPlayerBScore = 0 ' Resets the scores
blnPlayerTurn = True ' Makes it the turn of Player A
intCardCounter = 0 ' Sets the number of selected cards to 0
End Sub

Private Function DetermineGrid() As Integer
Dim intHeight As Integer
Select Case blnPairs
    Case Is = True
        Select Case intTotalNumber
            Case 6 To 10
                intHeight = 2
            Case 12 To 14
                intHeight = 3
            Case Is = 16
                intHeight = 4
            Case Is = 18
                intHeight = 3
            Case Is >= 20
                intHeight = 4
        End Select
    Case Is = False
        Select Case intTotalNumber
            Case Is <= 21
                intHeight = 3
            Case Is >= 24
                intHeight = 4
        End Select
End Select
DetermineGrid = intHeight
End Function

Private Sub GameType(blnTimeExpire As Boolean)
Select Case blnPairs
    Case Is = True
        If intCardCounter = 2 Then
            intTimePassed = 0 ' Resets the Timer
            intCardCounter = 0 'Resets the selected cards
            Call NextTurn ' Resets the game for the next turn
            Call SwitchTurn 'Changes Player
        End If
    Case Is = False
        If intCardCounter = 3 Then
            intTimePassed = 0 'Resets the Timer
            intCardCounter = 0
            Call NextTurn
            Call SwitchTurn
        End If
End Select

If blnTimeExpire = True Then ' As above
    intTimePassed = 0
    intCardCounter = 0
    Call NextTurn
    Call SwitchTurn
End If

Call UpdateBoxes
End Sub

Private Sub UpdateScore()
If blnPlayerTurn = True Then
    intPlayerAScore = intPlayerAScore + 1
        Else
    intPlayerBScore = intPlayerBScore + 1
End If
lblScore.Caption = "Player A: " & CStr(intPlayerAScore) & vbCrLf _
        & "Player B: " & CStr(intPlayerBScore)
End Sub

Private Sub UpdateBoxes()
If blnPlayerTurn = True Then
    lblInstruction.Caption = "Player A Turn"
Else
    lblInstruction.Caption = "Player B Turn"
End If
lblScore.Caption = "Player A: " & CStr(intPlayerAScore) & vbCrLf _
        & "Player B: " & CStr(intPlayerBScore)
If intTimeLimit <> 0 Then
    lblTimeLimit.Caption = "Time Remaining: " & intTimeLimit - intTimePassed
Else
    lblTimeLimit.Caption = "No Time Limit"
End If
End Sub

Private Sub SwitchTurn()
If blnPlayerTurn = True Then
    blnPlayerTurn = False
    lblInstruction.Caption = "Player B Turn"
Else
    blnPlayerTurn = True
    lblInstruction.Caption = "Player A Turn"
End If
End Sub

Private Sub mnuAbout_Click()
Pairs.Enabled = False
frmAbout.Show
frmAbout.Visible = True
End Sub

Private Sub mnuClose_Click()
Unload Me
End
End Sub

Private Sub mnuGameOptions_Click()
Me.Enabled = False
frmOptions.Show
frmOptions.Visible = True
End Sub

Private Sub mnuHowTo_Click()
Me.Enabled = False
frmInstructions.Visible = True
End Sub

Private Sub mnuNewGame_Click()
Call UnloadCards 'Gets rid of all the cards
Call InitialiseGame ' Starts the Game!
End Sub

Private Function RoundUp(sngRoundIt As Single) As Integer
If sngRoundIt = Int(sngRoundIt) Then RoundUp = sngRoundIt
If sngRoundIt <> Int(sngRoundIt) Then
    RoundUp = CInt(sngRoundIt + 0.5)
End If
End Function
