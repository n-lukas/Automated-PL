VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Tutorial 
   Caption         =   "Tutorial"
   ClientHeight    =   8250.001
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9390.001
   OleObjectBlob   =   "Tutorial.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Tutorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim page As Integer

Private Sub Label5_Click()

End Sub

Private Sub UserForm_Initialize()
    Tutorial.MultiPage1.Value = 0
End Sub

Private Sub CommandButton1_Click()
If page = 5 Then
    Tutorial.Hide
Else
    Tutorial.MultiPage1.Value = page + 1
    page = Tutorial.MultiPage1.Value
End If
End Sub


