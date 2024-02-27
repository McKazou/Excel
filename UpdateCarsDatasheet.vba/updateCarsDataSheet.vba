'This macro will get all URL in the table and update all external link then copy paste the value in the corresponding line in the table
Sub updateDataCars()


    Dim URLs As Range
    Set URLs = Range("Data_Cars[URL]")
    Debug.Print ("Start Looking in: " & (Range("Data_Cars[URL]").Row - 1) & " URL")
    Dim URLCell As Range
    Dim cpt As Integer
    cpt = 0
    For Each URLCell In URLs
        Debug.Print (URLCell.Value)
        cpt = cpt + 1
        If cpt <> 1 Then
        With Worksheet("CarsData")
            .Range("URL_Input").Value = URLCell.Value
            .Calculate
            Call updateDataconnections
            .Range("Data_Car_To_Copy").Copy _
                destination:=Range(Data_Car_To_Copy).Offset(cpt, 0)
        End If
    Next
    
End Sub

'This will update any connection in this workbook; in our case the HTML DATA query we have
Sub updateDataconnections()
    i = 1
    For Each wb_connection In ThisWorkbook.Connections
        Debug.Print wb_connection
        'Sheet1.Cells(i, 1).Value = wb_connection
        wb_connection.Refresh
        i = i + 1
    Next

End Sub
