Sub HighlightDuplicates()
    Dim ws As Worksheet
    Dim cell As Range
    Dim dict As Collection
    Dim lastRow As Long
    Dim textValue As String
    Dim i As Long
    
    ' Установите ссылку на активный лист
    Set ws = ThisWorkbook.Sheets("Реестр")
    
    ' Создаем коллекцию для хранения уникальных значений
    Set dict = New Collection
    
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    
    ' Сначала очищаем все предыдущие цвета
    ws.Range("D1:D" & lastRow).Interior.colorIndex = xlNone
    
    ' Проходим по всем ячейкам в столбце A и собираем информацию о дубликатах
    For Each cell In ws.Range("D1:D" & lastRow)
        textValue = cell.value
        
        If textValue <> "" Then ' Игнорируем пустые ячейки
            On Error Resume Next ' Игнорируем ошибки при добавлении дубликатов
            
            dict.Add textValue, CStr(textValue) ' Добавляем значение в коллекцию
            
            If Err.Number <> 0 Then
                ' Если значение уже существует (это дубликат), подкрашиваем текущую ячейку
                
                ' Подкрашиваем текущую ячейку и все дубликаты в зеленый цвет
                cell.Interior.Color = RGB(0, 255, 0) ' Зеленый цвет
                
                ' Подкрашиваем все найденные ранее ячейки с этим значением
                For i = 1 To lastRow
                    If ws.Cells(i, 1).value = textValue Then
                        ws.Cells(i, 1).Interior.Color = RGB(0, 255, 0) ' Зеленый цвет для всех дубликатов
                    End If
                Next i
                
            End If
            
            On Error GoTo 0 ' Включаем обработку ошибок снова
        End If
    Next cell
    
End Sub
