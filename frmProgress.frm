VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress 
   Caption         =   "Progress"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8685
   OleObjectBlob   =   "frmProgress.frx":0000
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
    subRemoveCloseButton Me
    Me.StartUpPosition = 0
    Me.Top = 150
    Me.Left = 150
    Call frmMain.GenerateTheFlimsy
End Sub

Private Sub UserForm_Initialize()
    lblBar.Width = 0
    lblBar.Caption = ""
    lblInfo.Caption = ""
End Sub

Public Sub CloseProgressBar()
    Unload frmProgress
End Sub
