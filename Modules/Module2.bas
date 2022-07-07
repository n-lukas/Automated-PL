Attribute VB_Name = "Module2"
Dim CellRef() As Variant
Dim CellLabel() As Variant

Sub BuildDashboard()
    Sheets.Add.Name = "Dashboard" 'Setting up Dashboard
    ActiveWindow.DisplayGridlines = False
    With ActiveSheet
        'Big Title
        With .Range("G5")
            .Value = "=CONCAT(""Dashboard for "",'Model'!C4)"
            .Font.Size = 48
        End With

        'Adding all labels
        CellRef = Array("C10", "D10", "C18", "D18", "K10", "U10", "T23")
        CellLabel = Array("Revenue", "% of Revenue", "Expenses", "% of Revenue", "Gross Profit Margin", "P&L Outlook", "Average Annual Growth Rate (AAGR)")
        For i = 0 To UBound(CellRef)
            With .Range(CellRef(i))
                .Interior.Color = RGB(0, 32, 96)
                .Font.Color = RGB(255, 255, 255)
                .Value = CellLabel(i)
            End With
        Next i

        'Adding Text Labels
        CellRef = Array("C11", "C12", "C13", "C19", "C20", "C21", "C22", "C23", "C24", "C25", "U11", "T24")
        CellLabel = Array("Sales", "Credit", "Other", "Cost of Sales", "SG&A", "Advertising", "R&D", "Fixed Cost", "Variable Cost", "Other", "Average Yearly EBITDA", "Revenue")
        For i = 0 To UBound(CellRef)
            With .Range(CellRef(i))
                .Value = CellLabel(i)
            End With
        Next i

        'Adding Formulas
        Dim Modelref As Variant
        CellRef = Array("D11", "D12", "D13", "D19", "D20", "D21", "D22", "D23", "D24", "D25", "L11", "M11", "V12", "U13", "T25")
        Modelref = Array("D10", "D11", "D12", "D15", "D19", "D21", "D22", "D23", "D24", "D25")
        ReDim CellLabel(15) As Variant

        'Revenue Formulas
        For i = 0 To 2
            CellLabel(i) = "=ABS(SUM('Model'!" & Modelref(i) & ":OFFSET('Model'!" & Modelref(i) & ",,Query!$L$5)))/SUM('Model'!$D$13:OFFSET('Model'!$D$13,,Query!$L$5))"
        Next i

        'Cost Formulas
        For i = 3 To 9
            CellLabel(i) = "=ABS(SUM('Model'!" & Modelref(i) & ":OFFSET('Model'!" & Modelref(i) & ",,Query!$L$5)))/SUM('Model'!$D$13:OFFSET('Model'!$D$13,,Query!$L$5))"
        Next i

        'Rest of the formulas
        CellLabel(10) = "=SUM('Model'!D16:OFFSET('Model'!D16,,Query!L5))/SUM('Model'!D13:OFFSET('Model'!D13,,Query!L5))" 'Profit Margin
        CellLabel(11) = "=100%-L11" 'Profit Margin Difference for Graph
        CellLabel(12) = "=AVERAGE('Model'!D34:OFFSET('Model'!D34,,Query!L5))" 'Average Yearly EBITDA Amount
        CellLabel(13) = "=IF(V12>0,""EBITDA is Positive"", ""EBITDA is Negative"")" 'EBITDA is positive or negative
        CellLabel(14) = "=AVERAGE('Model'!E45:OFFSET('Model'!E45,,Query!L5-1))" 'AAGR Revenue Growth %

        'Building Formulas
        For i = 0 To UBound(CellRef)
            With .Range(CellRef(i))
                .Value = CellLabel(i)
                If i < 12 Or i = 14 Then
                .NumberFormat = "0%"
                Else
                End If
            End With
        Next i

        'Cleaning up
        .Columns("C").ColumnWidth = 11.09
        .Columns("D").ColumnWidth = 11.45
        .Range("T24").Font.Bold = True
        .Range("L11").Font.Color = RGB(255, 255, 255)
        .Range("M11").Font.Color = RGB(255, 255, 255)
        .Range("V12").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

        If (ActiveSheet.Range("V12") > 0) Then
            .Range("U13").Font.Color = RGB(0, 176, 80)
        Else
            .Range("U13").Font.Color = RGB(255, 0, 0)
        End If

        'Add some borders
        CellRef = Array("C10:D13", "C18:D25", "U10:W18", "T23:X25")

        For Each item In CellRef
            With .Range(item)
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThick
                .Borders(xlEdgeRight).Weight = xlThick
                .Borders(xlEdgeLeft).Weight = xlThick
                .Borders(xlEdgeTop).Weight = xlThick
            End With
        Next item

        'Center align text
        CellRef = Array("K10:N10", "U10:W10", "T23:X23", "U11:W11", "U13:W13", "T24:X24", "T25:X25")
        For i = 0 To UBound(CellRef)
            With .Range(CellRef(i))
                .Merge
                .HorizontalAlignment = xlCenter
            End With
        Next i
        .Range("V12").HorizontalAlignment = xlCenter

        'Icon
        Attempts = 0
        Dim success As Boolean
        success = False
        Do Until success = True Or Attempts > 50 'Assumed problem with Excel, likes to fail paste function but should succeed in a subsequent try
            Attempts = Attempts + 1
            If (ActiveSheet.Range("V12") > 0) Then
                Worksheets("Validations").Range("D8:E12").CopyPicture (1)
            Else
                Worksheets("Validations").Range("D14:E18").CopyPicture (1)
            End If
            On Error GoTo Retry
                .Range("U14:W18").PasteSpecial (0)
                Selection.ShapeRange.IncrementLeft 20
                success = True
Retry:
        Loop

        'Graph
        Dim profitmargin As ChartObject
        Set profitmargin = Sheets("Dashboard").ChartObjects.Add(Left:=433, Width:=360, Top:=200, Height:=215) 'Left need adjusting
        With profitmargin.Chart
            .SetSourceData Source:=Sheets("Dashboard").Range("K10:O11")
            .ChartType = xlDoughnut
            .ChartStyle = 251
            .HasLegend = False
            .HasTitle = True
            .ChartColor = 19
            .ChartTitle.Text = "=Dashboard!$L$11"
            .ChartTitle.Font.Color = RGB(0, 0, 0)
            .ChartTitle.Font.Size = 28
            'Necessary Code block due to PlotArea function being known to fail
            With .PlotArea
                .Select
                With Selection
                    .Left = 105.09
                    .Top = 20.18
                    ActiveSheet.Range("A1").Activate
                End With
            End With
            'End necessary code block
            .ChartTitle.Left = 154.802
            .ChartTitle.Top = 73
            .SeriesCollection(2).Points(3).Interior.Color = RGB(195, 216, 187)
        End With

        .Range("A1").Select

    End With
End Sub
