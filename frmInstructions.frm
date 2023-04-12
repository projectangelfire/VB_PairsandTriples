VERSION 5.00
Begin VB.Form frmInstructions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "How to Play"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      Caption         =   $"frmInstructions.frx":0000
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmInstructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Call Form_Unload(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Pairs.Enabled = True
Unload Me
End Sub
