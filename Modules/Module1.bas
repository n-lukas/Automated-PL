Attribute VB_Name = "Module1"
Dim ModelYears As Integer
Dim StartYear As Integer
Dim CellRef(3) As String
Dim CurrentYear As String
Dim Labels As Variant


Sub BuildModel()

    ModelYears = ThisWorkbook.Sheets("Query").Range("L5").Value
    If (ThisWorkbook.Sheets("Settings").Range("D7").Value = "Default") Then
        StartYear = ThisWorkbook.Sheets("Query").Range("L3").Value
    Else
        StartYear = ThisWorkbook.Sheets("Settings").Range("D7").Value
    End If
    Sheets.Add.Name = "Model" 'Setting up Model
        ActiveWindow.DisplayGridlines = False
        With ActiveSheet
            .Columns("C").ColumnWidth = 25.64
            'Title Creation
            With .Range("C4")
                .Value = "=IF(COUNTA(Settings!D3)>0,Settings!D3,""Automated P&L"")"
                .Interior.Color = RGB(0, 32, 96)
                .Font.Color = RGB(255, 255, 255)
            End With
            With .Range("C5")
                .Value = "=CONCAT(""$ "", Settings!D13)"
                .Interior.Color = RGB(0, 32, 96)
                .Font.Color = RGB(255, 255, 255)
            End With
            'Initial Year Gray Block
            .Range("C7").Interior.Color = RGB(231, 230, 230)
            'Revenue Section
            .Range("C9").Font.Bold = True
            .Range("C9").Value = "Revenue"
            Labels = Array("Sales", "Credit", "Other")
            For i = 0 To 2
                .Range("C" & 10 + i).Value = Labels(i)
                .Range("C" & 10 + i).IndentLevel = 1
            Next i
            'Total Revenue
            With .Range("C13")
                .Font.Bold = True
                .Value = "Total Revenue"
                .IndentLevel = 1
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlMedium
            End With
            'Cost of Sales
            .Range("C15") = "Cost of Sales"
            'Gross Profit
            With .Range("C16")
                .Font.Bold = True
                .Value = "Gross Profit"
                .IndentLevel = 1
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlMedium
            End With
            'Expense section
            .Range("C18").Font.Bold = True
            .Range("C18").Value = "Expense"
            Labels = Array("SG&A", "Depreciation & Amortization", "Advertising", "R&D", "Fixed Cost", "Variable Cost", "Other")
            For i = 0 To 6
                .Range("C" & 19 + i).Value = Labels(i)
                .Range("C" & 19 + i).IndentLevel = 1
            Next i
            'Total Expenses
            With .Range("C26")
                .Font.Bold = True
                .Value = "Total Expenses"
                .IndentLevel = 1
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlMedium
            End With
            'EBIT
            With .Range("C28")
                .Font.Bold = True
                .Value = "EBIT"
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlMedium
                .Interior.Color = RGB(252, 228, 214)
            End With
            'Interest Expense
            .Range("C30").Value = "Interest Expense"
            .Range("C30").IndentLevel = 1
            'Taxes
            .Range("C31").Value = "Taxes"
            .Range("C31").IndentLevel = 1
            'Net Income
            With .Range("C32")
                .Font.Bold = True
                .Value = "Net Income"
                .IndentLevel = 1
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlMedium
            End With
            'EBITDA
            With .Range("C34")
                .Font.Bold = True
                .Value = "EBITDA"
                .Interior.Color = RGB(252, 228, 214)
            End With
            'Investment section
            .Range("C37").Font.Bold = True
            .Range("C37").Value = "Investment"
            'Taxes
            .Range("C38").Value = "Capex"
            .Range("C38").IndentLevel = 1
            'Other
            .Range("C39").Value = "Other"
            .Range("C39").IndentLevel = 1
            'Total Investment
            With .Range("C40")
                .Font.Bold = True
                .Value = "Total Investments"
                .IndentLevel = 1
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlMedium
            End With
            'Cumulative Investment
            .Range("C42").Value = "Cumulative Investment"
            .Range("C42").IndentLevel = 1
            'Cumulative Depreciation
            .Range("C43").Value = "Cumulative Depreciation"
            .Range("C43").IndentLevel = 1
            'Variance
            .Range("C44").Value = "Variance"
            .Range("C44").IndentLevel = 1
            'Revenue Growth
            .Range("C45").Value = "Revenue Growth"
            .Range("C45").IndentLevel = 1

            'Setting Up 2nd Line
            'Non Repetitions first
            .Range("D20").Value = "=IFERROR(SLN(D40,0,Validations!$C$45),0)" 'Depreciation & Amortization Part 1
            .Range("D20").Errors(4).Ignore = True 'Remove the pesky non-consistant formula error
            .Range("D42").Value = "=D40" 'Cumulative Investment Part 1
            'Repetitions
            For i = 0 To ModelYears
                'Title Block
                .Range("D4:D5").Offset(, i).Interior.Color = RGB(0, 32, 96)
                'Years
                .Range("D7").Offset(, i).Interior.Color = RGB(231, 230, 230)
                .Range("D7").Offset(, i).Value = StartYear + i
                'Sales
                CurrentYear = ActiveSheet.Range("D7").Offset(, i).Address 'Defines the current year for each loop. The loop will end at the last year of the model
                .Range("D10").Offset(, i).Value = "=IFERROR(GETPIVOTDATA(""Transformed Cost/Profit"",Query!$N$1,""Year""," & CurrentYear & ",""Transaction Type"",$C$9,""Sub Type"",$C10),0)" 'CurrentYear line changes year
                'Credit
                .Range("D11").Offset(, i).Value = "=IFERROR(GETPIVOTDATA(""Transformed Cost/Profit"",Query!$N$1,""Year""," & CurrentYear & ",""Transaction Type"",$C$9,""Sub Type"",$C11),0)"
                'Other
                .Range("D12").Offset(, i).Value = "=IFERROR(GETPIVOTDATA(""Transformed Cost/Profit"",Query!$N$1,""Year""," & CurrentYear & ",""Transaction Type"",$C$9,""Sub Type"",$C12),0)"
                'Total Revenue
                CellRef(0) = ActiveSheet.Range("D10").Offset(, i).Address 'Stores address of current cell based on year, used to change values within Sum function
                CellRef(1) = ActiveSheet.Range("D12").Offset(, i).Address
                With .Range("D13").Offset(, i)
                    .Value = "=SUM(" & CellRef(0) & ":" & CellRef(1) & ")"  'Uses above address to change reference
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).Weight = xlMedium
                End With
                'Cost of Sales
                .Range("D15").Offset(, i).Value = "=IFERROR(GETPIVOTDATA(""Transformed Cost/Profit"",Query!$N$1,""Year""," & CurrentYear & ",""Transaction Type"",$C$18,""Sub Type"",$C15),0)"
                'Gross Profit
                CellRef(0) = ActiveSheet.Range("D13").Offset(, i).Address
                CellRef(1) = ActiveSheet.Range("D15").Offset(, i).Address
                With .Range("D16").Offset(, i)
                    .Value = "=SUM(" & CellRef(0) & "," & CellRef(1) & ")" '
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).Weight = xlMedium
                End With
                'SG&A
                .Range("D19").Offset(, i).Value = "=IFERROR(GETPIVOTDATA(""Transformed Cost/Profit"",Query!$N$1,""Year""," & CurrentYear & ",""Transaction Type"",$C$18,""Sub Type"",$C19),0)"
                'Advertising
                .Range("D21").Offset(, i).Value = "=IFERROR(GETPIVOTDATA(""Transformed Cost/Profit"",Query!$N$1,""Year""," & CurrentYear & ",""Transaction Type"",$C$18,""Sub Type"",$C21),0)"
                'R&D
                .Range("D22").Offset(, i).Value = "=IFERROR(GETPIVOTDATA(""Transformed Cost/Profit"",Query!$N$1,""Year""," & CurrentYear & ",""Transaction Type"",$C$18,""Sub Type"",$C22),0)"
                'Fixed Cost
                .Range("D23").Offset(, i).Value = "=IFERROR(GETPIVOTDATA(""Transformed Cost/Profit"",Query!$N$1,""Year""," & CurrentYear & ",""Transaction Type"",$C$18,""Sub Type"",$C23),0)"
                'Variable Cost
                .Range("D24").Offset(, i).Value = "=IFERROR(GETPIVOTDATA(""Transformed Cost/Profit"",Query!$N$1,""Year""," & CurrentYear & ",""Transaction Type"",$C$18,""Sub Type"",$C24),0)"
                'Other
                .Range("D25").Offset(, i).Value = "=IFERROR(GETPIVOTDATA(""Transformed Cost/Profit"",Query!$N$1,""Year""," & CurrentYear & ",""Transaction Type"",$C$18,""Sub Type"",$C25),0)"
                'Total Expense
                CellRef(0) = ActiveSheet.Range("D19").Offset(, i).Address 'Stores address of current cell based on year, used to change values within Sum function
                CellRef(1) = ActiveSheet.Range("D25").Offset(, i).Address
                With .Range("D26").Offset(, i)
                    .Value = "=SUM(" & CellRef(0) & ":" & CellRef(1) & ")"  'Uses above address to change reference
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).Weight = xlMedium
                End With
                'EBIT
                CellRef(0) = ActiveSheet.Range("D16").Offset(, i).Address 'Stores address of current cell based on year, used to change values within Sum function
                CellRef(1) = ActiveSheet.Range("D26").Offset(, i).Address
                With .Range("D28").Offset(, i)
                    .Value = "=SUM(" & CellRef(0) & "," & CellRef(1) & ")"  'Uses above address to change reference
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).Weight = xlMedium
                    .Interior.Color = RGB(252, 228, 214)
                End With
                'Interest Expense
                .Range("D30").Offset(, i).Value = "=IFERROR(GETPIVOTDATA(""Transformed Cost/Profit"",Query!$N$1,""Year""," & CurrentYear & ",""Sub Type"",$C30),0)"
                'Taxes
                CellRef(0) = ActiveSheet.Range("D28").Offset(, i).Address
                CellRef(1) = ActiveSheet.Range("D30").Offset(, i).Address
                .Range("D31").Offset(, i).Value = "=MIN(-(" & CellRef(0) & "+" & CellRef(1) & ")*(Settings!$D$11/100),0)"
                'Net Income
                CellRef(0) = ActiveSheet.Range("D28").Offset(, i).Address 'Stores address of current cell based on year, used to change values within Sum function
                CellRef(1) = ActiveSheet.Range("D30").Offset(, i).Address
                CellRef(2) = ActiveSheet.Range("D31").Offset(, i).Address
                With .Range("D32").Offset(, i)
                    .Value = "=SUM(" & CellRef(0) & "," & CellRef(1) & "," & CellRef(2) & ")"  'Uses above address to change reference
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).Weight = xlMedium
                End With
                'EBITDA
                CellRef(0) = ActiveSheet.Range("D28").Offset(, i).Address 'Stores address of current cell based on year, used to change values within Sum function
                CellRef(1) = ActiveSheet.Range("D20").Offset(, i).Address
                With .Range("D34").Offset(, i)
                    .Value = "=SUM(" & CellRef(0) & ",-" & CellRef(1) & ")"  'Uses above address to change reference
                    .Interior.Color = RGB(252, 228, 214)
                End With
                'Capex
                .Range("D38").Offset(, i).Value = "=IFERROR(GETPIVOTDATA(""Transformed Cost/Profit"",Query!$N$1,""Year""," & CurrentYear & ",""Transaction Type"",$C$37,""Sub Type"",$C38),0)"
                'Other
                .Range("D39").Offset(, i).Value = "=IFERROR(GETPIVOTDATA(""Transformed Cost/Profit"",Query!$N$1,""Year""," & CurrentYear & ",""Transaction Type"",$C$37,""Sub Type"",$C39),0)"
                'Total Investment
                CellRef(0) = ActiveSheet.Range("D38").Offset(, i).Address 'Stores address of current cell based on year, used to change values within Sum function
                CellRef(1) = ActiveSheet.Range("D39").Offset(, i).Address
                With .Range("D40").Offset(, i)
                    .Value = "=SUM(" & CellRef(0) & ":" & CellRef(1) & ")"  'Uses above address to change reference
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).Weight = xlMedium
                End With

                'Format all to accounting
                Labels = Array("D10:D13", "D15:D16", "D19:D26", "D28", "D30:D32", "D34", "D38:D40")
                Dim item As Variant
                For Each item In Labels
                    .Range(item).Offset(, i).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Next item
                'Setting up 3rd and final line repetitions

                If (i <= ModelYears - 1) Then 'Since this is in column 3, this makes sure it doesn't get double added
                    'Depreciation & Amortization Part 2
                    CellRef(0) = ActiveSheet.Range("E44").Offset(, i).Address 'Stores address of current cell based on year, used to change values within Sum function
                    CellRef(1) = ActiveSheet.Range("D20").Offset(, i).Address
                    CellRef(2) = ActiveSheet.Range("E40").Offset(, i).Address
                    .Range("E20").Offset(, i).Value = "=IFERROR(IF(" & CellRef(0) & "<0,IF(" & CellRef(1) & "+SLN(" & CellRef(2) & ",0,Validations!$C$45)<" & CellRef(0) & "," & CellRef(0) & "," & CellRef(1) & "+SLN(" & CellRef(2) & ",0,Validations!$C$45)),SLN(" & CellRef(2) & ",0,Validations!$C$45)),0)"
                    .Range("E20").Offset(, i).Errors(4).Ignore = True
                    'Cumulative Investment Part 2
                    CellRef(0) = ActiveSheet.Range("E40").Offset(, i).Address
                    CellRef(1) = ActiveSheet.Range("D42").Offset(, i).Address
                    .Range("E42").Offset(, i).Value = "=SUM(" & CellRef(0) & "+" & CellRef(1) & ")"
                    'Cumulative Depreciation
                    CellRef(0) = ActiveSheet.Range("D20").Offset(, i).Address
                    .Range("E43").Offset(, i).Value = "=SUM($D$20:" & CellRef(0) & ")"
                    'Variance
                    CellRef(0) = ActiveSheet.Range("E42").Offset(, i).Address
                    CellRef(1) = ActiveSheet.Range("E43").Offset(, i).Address
                    .Range("E44").Offset(, i).Value = "=" & CellRef(0) & "-" & CellRef(1)
                    'Revenue Growth
                    CellRef(0) = ActiveSheet.Range("E13").Offset(, i).Address
                    CellRef(1) = ActiveSheet.Range("D13").Offset(, i).Address
                    .Range("E45").Offset(, i).Value = "=IFERROR((" & CellRef(0) & "-" & CellRef(1) & ")/" & CellRef(1) & ",0)"
                Else
                End If
            Next i

        'Model Fully Built

        ' Hide rows 42-45
        .Range("A42:A45").EntireRow.Hidden = True

        End With
            
    Application.Run "Module2.BuildDashboard"

End Sub


