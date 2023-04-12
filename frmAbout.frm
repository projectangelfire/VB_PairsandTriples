VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Pairs"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Call Form_Unload(0)
End Sub

Private Sub Form_Load()
lblAbout.Caption = "Pairs + Trebles!" & vbCrLf & _
    "Designed for Applied Computing" & vbCrLf & _
    "Copyright(c), 2006" & vbCrLf & _
    "By Andy Ball"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Pairs.Enabled = True
Me.Visible = False
Unload Me
End Sub

