VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ticker_names()
Dim column As Integer
Dim row As Integer
row = 2
column = 1
    For i = 2 To 705714
            If Cells(i + 1, column).Value <> Cells(i, column).Value Then
                Cells(row, 8).Value = Cells(i, column).Value
                row = row + 1
            End If
    Next
End Sub
Sub yearlychange()
    Dim foundTicker As Boolean
    Dim yearly_change As Double
    Dim opening_price As Double
    Dim closing_price As Double
        foundTicker = False
        For i = 2 To 705714
            If foundTicker = False Then
                ticker_symbol = Cells(i, 1).Value
                opening_price = Cells(i, 3).Value
                foundTicker = True
            End If
            If Cells(i + 1, 1).Value <> ticker_symbol Then
                closing_price = Cells(i, 6).Value
                yearly_change = closing_price - opening_price
                For j = 2 To 2836
                    If Cells(j, 8).Value = ticker_symbol Then
                       Cells(j, 9).Value = yearly_change
                       foundTicker = False
                    End If
                Next
            End If
        Next
End Sub
Sub percent_change_function()
Dim percent_change As Double
Dim yearly_change As Double
Dim opening_price As Double
Dim foundTicker As Boolean
foundTicker = False
    For i = 2 To 705714
        If foundTicker = False Then
            ticker_symbol = Cells(i, 1).Value
            opening_price = Cells(i, 3).Value
            foundTicker = True
        End If
        If Cells(i + 1, 1).Value <> ticker_symbol Then
            For j = 2 To 2836
                If Cells(j, 8).Value = ticker_symbol Then
                    yearly_change = Cells(j, 9).Value
                    percent_change = (yearly_change / opening_price) * 100
                    Cells(j, 10).Value = Str(percent_change) + "%"
                    foundTicker = False
                End If
              Next
        End If
    Next
End Sub
Sub total_stock_volume()
    Dim total_volume As Variant
    Dim individual_volume As Long
    total_volume = 0
    j = 2
    For i = 2 To 705714
            If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
                individual_volume = Cells(i, 7).Value
                total_volume = total_volume + individual_volume
            ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                    total_volume = total_volume + Cells(i, 7).Value
                    Cells(j, 11).Value = total_volume
                    j = j + 1
                    total_volume = 0
            End If
    Next i
End Sub


Sub greatest()
Dim max_percent As Double
Dim min_percent As Double
Dim max_volume As LongLong
max_percent = WorksheetFunction.Max(Range("J2:J290"))
Cells(2, 15).Value = Str(max_percent * 100) + "%"
min_percent = WorksheetFunction.Min(Range("J2:J290"))
Cells(3, 15).Value = Str(min_percent * 100) + "%"
max_volume = WorksheetFunction.Max(Range("K2:K290"))
Cells(4, 15).Value = max_volume

For j = 2 To 2836
    If Cells(j, 10).Value = max_percent Then
        Cells(2, 14).Value = Cells(j, 8).Value
    ElseIf Cells(j, 10).Value = min_percent Then
        Cells(3, 14).Value = Cells(j, 8).Value
    ElseIf Cells(j, 11).Value = max_volume Then
        Cells(4, 14).Value = Cells(j, 8).Value
    End If
Next
End Sub
    


