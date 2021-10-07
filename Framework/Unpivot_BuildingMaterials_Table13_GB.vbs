Sub Unpivot()

Dim a, b, i As Long, ii As Long, n As Long, temp
Set w1 = Sheets("temp")
Application.DisplayAlerts = False 'switching off the alert button
Application.ScreenUpdating = False
    
    'name of sheet is Sheet1
    a = Sheets("temp").Cells(1).CurrentRegion.Value
    
    '11 is the number of columns after unpivot
    ReDim b(1 To UBound(a, 1) * UBound(a, 2), 1 To 6)
    For i = 1 To UBound(a, 1)
        '9 is the last column to unpivot by, country/series specific notes is the column name
        If a(i, 4) = "Other Building Material Type" Then
            temp = i
        Else
            If Application.CountA(Application.Index(a, i, 0)) > 2 Then
                '11-1=10
                For ii = 5 To UBound(a, 2)
                    If a(i, ii) <> "" Then
                        n = n + 1
                         b(n, 1) = a(i, 1): b(n, 2) = a(i, 2)
                         b(n, 1) = a(i, 1): b(n, 3) = a(i, 3)
                         b(n, 1) = a(i, 1): b(n, 4) = a(i, 4)
                         b(n, 1) = a(i, 1): b(n, 5) = a(i, 5)
                         b(n, 5) = a(temp, ii): b(n, 6) = a(i, ii)
                              
                    End If
                Next
            End If
        End If
    Next
    
'
    With Sheets.Add.Cells(1).CurrentRegion.Resize(, UBound(b, 2))
        .Value = [{"Date","Geographical Region","Building Material Type","Other Building Material Type","Production/Deliveries/Stocks","Th.sq.m. of roof area covered"}]
        .Rows(2).Resize(n).Value = b
        .EntireColumn.AutoFit
    End With


Application.DisplayAlerts = True 'switching on the alert button
Application.ScreenUpdating = True

End Sub
    






