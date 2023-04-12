VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   5145
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraGameType 
      Caption         =   "Game Type"
      Height          =   855
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   2415
      Begin VB.OptionButton btnPairs 
         Caption         =   "Option1"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   255
      End
      Begin VB.OptionButton btnTrebles 
         Caption         =   "Option2"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblPairs 
         Caption         =   "Pairs"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblTrebles 
         Caption         =   "Trebles"
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit No Save"
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Exit Saving Options"
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton btnDefaults 
      Caption         =   "Load Defaults"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   4080
      Width           =   1335
   End
   Begin VB.ComboBox cboCardGraphic 
      Height          =   315
      ItemData        =   "frmOptions.frx":0000
      Left            =   240
      List            =   "frmOptions.frx":000D
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
   End
   Begin VB.ComboBox cboBackground 
      Height          =   315
      ItemData        =   "frmOptions.frx":002F
      Left            =   240
      List            =   "frmOptions.frx":0054
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txtTimeLimit 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtNoPairs 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblCardGraphic 
      Caption         =   "Cards Theme"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label lblBackground 
      Caption         =   "Background"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label lblTimeLimit 
      Caption         =   "Time Limit (2-12s) 0 = No Time Limit"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label lblNoPairs 
      Caption         =   "Number of Pairs (5-15)"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetShortPathName Lib "kernel32" Alias _
"GetShortPathNameA" (ByVal lpszLongPath As String, _
ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Sub btnDefaults_Click()
Dim intNoPairs As Integer
Dim blnPairs As Boolean
Dim intTimeLimit As Integer
Dim strBackground As String
Dim strCardGraphic As String

intNoPairs = 5
blnPairs = True
intTimeLimit = 0
strBackground = "StarGate"
strCardGraphic = "StarGate"

txtNoPairs.Text = CStr(intNoPairs)
btnPairs.Value = CBool(blnPairs)
txtTimeLimit.Text = CStr(intTimeLimit)
cboBackground.Text = CStr(strBackground)
cboCardGraphic.Text = CStr(strCardGraphic)

On Error GoTo 2 ' Terrible code but what choice do I have?

Dim intFree As Integer
intFree = FreeFile() 'Slightly annoying to do this but it avoids problems
Open GetShortName(CurDir() & "\options.dat") For Random As #intFree
    Put #intFree, 1, intNoPairs
    Put #intFree, 2, blnPairs
    Put #intFree, 3, intTimeLimit
    Put #intFree, 4, strBackground
    Put #intFree, 5, strCardGraphic
Close #intFree
Exit Sub ' Nobel prize for most Microsoft like program

2 MsgBox ("This application was designed to be run from a hard drive, " & vbCrLf & _
        "Although it can be run from a CD, the options are not alterable at this time")
End Sub

Private Sub btnSave_Click()
If CheckforErrors = True Then Exit Sub
On Error GoTo 1

Dim intFree As Integer
    intFree = FreeFile()
Open GetShortName(CurDir() & "\options.dat") For Random As #intFree
    Put #intFree, 1, CInt(txtNoPairs.Text)
    Put #intFree, 2, CBool(btnPairs.Value)
    Put #intFree, 3, CInt(txtTimeLimit.Text)
    Put #intFree, 4, CStr(cboBackground.Text)
    Put #intFree, 5, CStr(cboCardGraphic.Text)
Close #intFree
Dim strPicture As String
strPicture = CStr(cboBackground.Text)
If CheckExist(GetShortName(CurDir() & "\graphics\BackDrop\" & _
    strPicture & ".jpg")) = True Then
Pairs.Picture = LoadPicture(GetShortName(CurDir() & "\graphics\BackDrop\" & _
    strPicture & ".jpg"))
Dim intTemp As Integer
    intTemp = MsgBox("You will need to restart the Game for changes to " & _
        "take effect", vbOKOnly, "Notice!")
End If
Unload Me
Exit Sub
1 MsgBox ("This application was designed to be run from a hard drive, " & vbCrLf & _
        "Although it can be run from a CD, the options are not alterable at this time")
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim NoPairs As Integer
Dim Pairs As Boolean
Dim TimeLimit As Integer
Dim Background As String
Dim CardBack As String
Dim CardGraphic As String

Dim intFree As Integer
    intFree = FreeFile()
Open GetShortName(CurDir() & "\options.dat") For Random As #intFree
    Get #intFree, 1, NoPairs
    Get #intFree, 2, Pairs
    Get #intFree, 3, TimeLimit
    Get #intFree, 4, Background
    Get #intFree, 5, CardGraphic
Close #intFree

txtNoPairs.Text = CStr(NoPairs)
btnPairs.Value = CBool(Pairs)
    If btnPairs.Value = False Then btnTrebles.Value = True
txtTimeLimit.Text = CStr(TimeLimit)
cboBackground.Text = Background
cboCardGraphic.Text = CardGraphic

End Sub

Private Sub Form_Unload(Cancel As Integer)
Pairs.Enabled = True
End Sub

Private Sub txtNoPairs_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If Len(txtNoPairs.Text) > 1 Then KeyAscii = 0
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtTimeLimit_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If Len(txtTimeLimit.Text) > 1 Then KeyAscii = 0
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub cboBackground_KeyPress(KeyAscii As Integer)
KeyAscii = 0 ' Prevents any ascii character from being entered
End Sub

Private Sub cboCardGraphic_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Function CheckforErrors() As Boolean
Dim intTemp As Integer

If CInt(txtNoPairs.Text) < 3 Then
    intTemp = MsgBox("The number of Pairs must be three or greater!", vbOKOnly, "Error!")
    txtNoPairs.Text = "3"
    CheckforErrors = True
ElseIf CInt(txtNoPairs.Text) > 14 And btnPairs.Value = True Then
    intTemp = MsgBox("The number of Pairs must be less than 14", vbOKOnly, "Error!")
    txtNoPairs.Text = "14"
    CheckforErrors = True
ElseIf CInt(txtNoPairs.Text) > 9 And btnTrebles.Value = True Then
    intTemp = MsgBox("The number of trebles must be less than 9", vbOKOnly, "Error!")
    txtNoPairs.Text = "9"
    CheckforErrors = True
ElseIf IsNumeric(CInt(txtNoPairs.Text)) = False Then
    intTemp = MsgBox("Please enter only numeric characters!", vbOKOnly, "Type Mismatch!")
    txtNoPairs.Text = "5"
    CheckforErrors = True
End If

If CInt(txtTimeLimit.Text) < 2 And CInt(txtTimeLimit.Text) <> 0 Then
    intTemp = MsgBox("Timer Value must be greater than 2 seconds", vbOKOnly, "Error!")
    CheckforErrors = True
ElseIf CInt(txtTimeLimit.Text) > 12 Then
    intTemp = MsgBox("Time Limit must be less than 12 seconds", vbOKOnly, "Error!")
    CheckforErrors = True
ElseIf IsNumeric(CInt(txtTimeLimit.Text)) = False Then
    intTemp = MsgBox("Please enter only Numeric characters", vbOKOnly, "Type Mismatch!")
    CheckforErrors = True
End If

Dim intCounter As Integer, intChecker As Integer
' Checks the cboBox for errors that will usually only occur if the user is
' deliberately trying to crash the program
For intCounter = 0 To cboBackground.ListCount - 1
    If cboBackground.Text <> cboBackground.List(intCounter) Then
        intChecker = intChecker + 1
    End If
Next intCounter
If intChecker = cboBackground.ListCount Then
    intTemp = MsgBox("Please Select a valid background", vbOKOnly, "Error!")
    cboBackground.Text = "Dreylor"
    CheckforErrors = True
End If

' Same routine as above except it checks the Card Theme comboBox
intChecker = 0: intCounter = 0
For intCounter = 0 To cboCardGraphic.ListCount - 1
    If cboCardGraphic.Text <> cboCardGraphic.List(intCounter) Then
        intChecker = intChecker + 1
    End If
Next intCounter
If intChecker = cboCardGraphic.ListCount Then
    intTemp = MsgBox("Please Select a valid theme", vbOKOnly, "Error!")
    cboCardGraphic.Text = "StarTrek"
    CheckforErrors = True
End If
End Function

Public Function CheckExist(strFile As String) As Boolean
If Dir$(strFile) <> "" Then
    CheckExist = True
Else
    CheckExist = False
    MsgBox ("A file required by pairs appears to have become corrupted " & vbCrLf & _
        "or missing. Please reinstall the application.")
End If
End Function

Function GetShortName(ByVal sLongFileName As String) As String
Dim lRetVal As Long, sShortPathName As String, iLen As Integer
'Buffer area for API function call return
sShortPathName = Space(255)
iLen = Len(sShortPathName)
lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
GetShortName = Left(sShortPathName, lRetVal)
End Function



