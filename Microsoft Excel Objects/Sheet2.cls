VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim BuildModel As Boolean

Private Sub CommandButton1_Click()
    UserForm1.Show
    With ThisWorkbook.Sheets("Settings")
        .Range("D3").Value = UserForm1.TextBox1.Value
        .Range("D5").Value = UserForm1.ListBox1.Value
        .Range("D7").Value = UserForm1.ListBox2.Value
        .Range("D9").Value = UserForm1.ListBox3.Value
        .Range("D11").Value = UserForm1.TextBox2.Value
        .Range("D13").Value = UserForm1.ComboBox1.Value
    End With
    If ThisWorkbook.Sheets("Settings").Range("D5").Value = "" Then
        ThisWorkbook.Sheets("Settings").Range("D5").Value = "Lifecycle"
    Else
    End If
    If ThisWorkbook.Sheets("Settings").Range("D7").Value = "" Then
        ThisWorkbook.Sheets("Settings").Range("D7").Value = "Default"
    Else
    End If
    If ThisWorkbook.Sheets("Settings").Range("D9").Value = "" Then
        ThisWorkbook.Sheets("Settings").Range("D9").Value = "Default"
    Else
    End If
    If ThisWorkbook.Sheets("Settings").Range("D13").Value = "" Then
    ThisWorkbook.Sheets("Settings").Range("D13").Value = "Actual Value"
    Else
    End If
    Unload UserForm1
End Sub

Private Sub CommandButton2_Click()
    Tutorial.Show
    Unload Tutorial
End Sub

Private Sub CommandButton3_Click()
    
   Dim connectnum As Long
    With ActiveWorkbook
        For connectnum = 1 To .Connections.Count
          If .Connections(connectnum).Type = xlConnectionTypeOLEDB Then
            .Connections(connectnum).OLEDBConnection.BackgroundQuery = False 'Pauses background refresh so tables get updated before continuing code
            ActiveWorkbook.RefreshAll
            .Connections(connectnum).OLEDBConnection.BackgroundQuery = True 'Turns back on background refresh and continues with code
          End If
        Next connectnum
    End With
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
   
    If BuildModel = False Then
        Application.Run "Module1.BuildModel"
    Else
        ThisWorkbook.Sheets("Model").Delete
        ThisWorkbook.Sheets("Dashboard").Delete
        Application.Run "Module1.BuildModel"
    End If

    BuildModel = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Private Sub CommandButton4_Click()
    Application.Run "Module3.AddRow"
End Sub
