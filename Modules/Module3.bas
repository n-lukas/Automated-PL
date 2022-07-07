Attribute VB_Name = "Module3"
Sub AddRow()

Dim TransactionTable As ListObject
Set TransactionTable = ActiveSheet.ListObjects("Table1")
TransactionTable.ListRows.Add

End Sub


