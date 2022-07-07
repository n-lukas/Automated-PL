VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Settings"
   ClientHeight    =   6250
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6260
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ListBox1.RowSource = "Validations!B29:B44"
    ListBox2.RowSource = "Validations!B48:B148"
    ListBox3.RowSource = "Validations!B48:B148"
    ComboBox1.RowSource = "Validations!D3:D5"
    With UserForm1
            .TextBox1.Value = ThisWorkbook.Sheets("Settings").Range("D3").Value
            .ListBox1.Value = ThisWorkbook.Sheets("Settings").Range("D5").Value
            .ListBox2.Value = ThisWorkbook.Sheets("Settings").Range("D7").Value
            .ListBox3.Value = ThisWorkbook.Sheets("Settings").Range("D9").Value
            .TextBox2.Value = ThisWorkbook.Sheets("Settings").Range("D11").Value
            .ComboBox1.Value = ThisWorkbook.Sheets("Settings").Range("D13").Value
    End With
End Sub

Private Sub CommandButton1_Click()
    If IsNumeric(UserForm1.TextBox2.Value) = True Then
        UserForm1.Hide
    Else
        MsgBox "Tax Rate must be a number", vbOkayOnly + vbCritical, "Error"
    End If
End Sub

Private Sub CommandButton2_Click()
    MsgBox "Name of P&L - Name you want on the P&L for your reference" & vbCrLf & vbCrLf & "Years to Amortize Over - Years to amortize the capital over. Lifecycle is over the entire P&L life" & vbCrLf & vbCrLf & "Start Year - First year of P&L. Default is the year of the earliest transaction" & vbCrLf & vbCrLf & "End Year - Last year of P&L. Default is the year of the latest transaction" & vbCrLf & vbCrLf & "Tax Rate - Corporate Tax Rate. Default is 21%", vbOkayOnly + vbInformation, "Help"
End Sub



