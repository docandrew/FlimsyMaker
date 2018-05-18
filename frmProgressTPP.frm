VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgressTPP 
   Caption         =   "Progress"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8625
   OleObjectBlob   =   "frmProgressTPP.frx":0000
End
Attribute VB_Name = "frmProgressTPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fileToGet As String
Private placeToStore As String

Private Sub UserForm_Activate()
    subRemoveCloseButton Me
    Me.StartUpPosition = 0
    Me.Top = 150
    Me.Left = 150
    Call DownloadBigFile(fileToGet, placeToStore, True)
End Sub

Private Sub UserForm_Initialize()
    lblBar.Width = 0
    lblBar.Caption = ""
    lblInfo.Caption = ""
End Sub

Public Sub CloseProgressBarTPP()
    Unload frmProgressTPP
End Sub

Property Let DownloadSource(src As String)
    fileToGet = src
End Property

Property Let StorageLocation(spot As String)
    placeToStore = spot
End Property
