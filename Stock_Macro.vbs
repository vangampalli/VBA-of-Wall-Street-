Attribute VB_Name = "Module1"
Sub stock()
    
    Dim New_List As Integer
    Dim Two_List As Integer
    Dim Price_Count As Integer
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Volume As Variant
    Dim Vol_count As Integer
    
'Get last row
    Dim Last_row As Long
    Last_row = Cells(Rows.Count, 1).End(xlUp).Row
    

'The ticker symbol.
    Cells(1, 9) = "Ticker"
    For i = 1 To Last_row
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            New_List = New_List + 1
            Cells(New_List + 1, 9) = Cells(i + 1, 1).Value
        End If
    Next i
    
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'Total Stock Volume

    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
    For i = 2 To Last_row
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            Open_Price = Cells(i, 3).Value
            Price_Count = Price_Count + 1
            'MsgBox (Open_Price)
        End If
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Close_Price = Cells(i, 6).Value
            'MsgBox (Close_Price)
        End If
        Cells(Price_Count + 1, 10).Value = Close_Price - Open_Price
        If Open_Price = 0 And Close_Price = 0 Then
            Cells(Price_Count + 1, 11) = Round(0, 2)
        ElseIf Open_Price = 0 Then
            Cells(Price_Count + 1, 11) = Round(0, 2)
        Else
            Cells(Price_Count + 1, 11).Value = Round(((Close_Price - Open_Price) / (Open_Price) * 100), 2)
        End If
        If Cells(Price_Count + 1, 10).Value >= 0 Then
                    Cells(Price_Count + 1, 10).Interior.ColorIndex = 4
                Else
                    Cells(Price_Count + 1, 10).Interior.ColorIndex = 3
                End If
        If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
            Volume = Volume + Cells(i, 7).Value
            Cells(Price_Count + 1, 12).Value = Volume
        End If
    Next i
    
End Sub

