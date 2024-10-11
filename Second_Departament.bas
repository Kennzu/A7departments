Attribute VB_Name = "Module2"
Sub РаспределитьЗаявкиДокумент()
    Dim wsСводная As Worksheet, wsТабель As Worksheet
    Dim lastRowСводная As Long, lastRowТабель As Long, lastColТабель As Long
    Dim выбраннаяДата As Variant
    Dim присутствующие() As String
    Dim countПрисутствующие As Long
    Dim распределеноЗаявок As Long
    Dim ФИО As String
    Dim новыйЛист As Worksheet
    Dim безопасноеИмя As String
    Dim i As Long, j As Long, k As Long
    Dim lastProcessedEmployeeIndex As Long
    
    ' Установка листов
    Set wsСводная = ThisWorkbook.Sheets("СВОДНАЯ")
    Set wsТабель = ThisWorkbook.Sheets("Табель документооборот")

    ' Считываем индекс последнего обработанного сотрудника из ячейки (например, AY2)
    lastProcessedEmployeeIndex = wsСводная.Range("AY2").Value

    ' Запрос даты у пользователя
    выбраннаяДата = InputBox("Введите дату для распределения заявок (в формате ДД.ММ.ГГГГ):", "Выбор даты")
    If Not IsDate(выбраннаяДата) Then
        MsgBox "Операция отменена пользователем или введена некорректная дата.", vbInformation
        Exit Sub
    End If
    выбраннаяДата = DateValue(выбраннаяДата)

    ' Определение последних строк и столбцов
    lastRowСводная = wsСводная.Cells(wsСводная.Rows.count, "A").End(xlUp).Row
    lastRowТабель = wsТабель.Cells(wsТабель.Rows.count, "B").End(xlUp).Row
    lastColТабель = wsТабель.Cells(1, wsТабель.Columns.count).End(xlToLeft).Column

    ' Собрать список присутствующих сотрудников
    ReDim присутствующие(1 To lastRowТабель - 1)
    countПрисутствующие = 0

    For j = 3 To lastColТабель ' Начинаем с C1
        If IsDate(wsТабель.Cells(1, j).Value) And DateValue(wsТабель.Cells(1, j).Value) = выбраннаяДата Then
            ' Собрать список присутствующих сотрудников
            For k = 2 To lastRowТабель
                If wsТабель.Cells(k, j).Value = "Да" Then
                    countПрисутствующие = countПрисутствующие + 1
                    присутствующие(countПрисутствующие) = wsТабель.Cells(k, "B").Value ' ФИО из столбца B

                    ' Создаем безопасное имя для листа (по руководителю)
                    безопасноеИмя = Application.WorksheetFunction.Substitute(wsТабель.Cells(k, "A").Value, "/", " ")

                    ' Проверяем, существует ли лист, если нет, создаем его
                    On Error Resume Next
                    Set новыйЛист = ThisWorkbook.Sheets(безопасноеИмя)
                    On Error GoTo 0

                    ' Если лист не существует, создаем его
                    If новыйЛист Is Nothing Then
                        Set новыйЛист = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
                        новыйЛист.Name = безопасноеИмя
                    End If

                    ' Сброс ссылки на новый лист для следующей итерации
                    Set новыйЛист = Nothing
                End If
            Next k
            Exit For ' Выходим из цикла после нахождения даты
        End If
    Next j

    If countПрисутствующие = 0 Then
        MsgBox "На выбранную дату нет присутствующих сотрудников.", vbExclamation
        Exit Sub
    End If

    ' Создаем коллекцию для хранения уже распределенных номеров заявок и соответствующих сотрудников
    Dim distributedRequests As Collection
    Set distributedRequests = New Collection

    ' Заполняем коллекцию уже распределенными номерами заявок
    For i = 4 To lastRowСводная
        If wsСводная.Cells(i, "T").Value <> "" Then ' Если ячейка T не пустая
            distributedRequests.Add wsСводная.Cells(i, 2).Value & "|" & wsСводная.Cells(i, "T").Value ' Номер заявки и ФИО сотрудника
        End If
    Next i

    ' Распределить заявки на эти листы начиная с последнего обработанного сотрудника + 1
    Dim currentСотрудник As Long
    currentСотрудник = lastProcessedEmployeeIndex + 1

    For i = 4 To lastRowСводная ' Начинаем с A4 для заявок
        Dim датаЗаявки As Variant
        датаЗаявки = wsСводная.Cells(i, 1).Value ' Дата заявки из столбца A
        
        If IsDate(датаЗаявки) Then
            ' Сравниваем только даты без учета времени и проверяем номер заявки
            If DateValue(датаЗаявки) = выбраннаяДата Then
                
                ' Проверка на наличие номера заявки в коллекции
                Dim requestExists As Boolean
                requestExists = False

                For Each Item In distributedRequests
                    If InStr(Item, wsСводная.Cells(i, 2).Value) > 0 Then
                        requestExists = True
                        Exit For
                    End If
                Next Item

                If Not requestExists Then
                    ' Проверяем, есть ли уже запись в столбце T для этой заявки
                    If wsСводная.Cells(i, "T").Value = "" Then ' Если ячейка пустая, значит не распределена
                        ' Получаем ФИО сотрудника
                        ФИО = присутствующие(currentСотрудник)
                        ' Записываем данные в соответствующие ячейки
                        wsСводная.Cells(i, "S").Value = wsТабель.Cells(Application.Match(ФИО, wsТабель.Range("B:B"), 0), "A").Value ' Руководитель
                        wsСводная.Cells(i, "T").Value = ФИО ' Сотрудник

                        ' Проверка на наличие заявки на листе руководителя
                        безопасноеИмя = Application.WorksheetFunction.Substitute(wsТабель.Cells(Application.Match(ФИО, wsТабель.Range("B:B"), 0), "A").Value, "/", " ")
                        
                        On Error Resume Next
                        Set новыйЛист = ThisWorkbook.Sheets(безопасноеИмя)
                        On Error GoTo 0

                        Dim заявкаУжеРаспределена As Boolean
                        заявкаУжеРаспределена = False

                        ' Проверяем, существует ли заявка на листе руководителя
                        If Not новыйЛист Is Nothing Then
                            Dim lastRowNew As Long
                            lastRowNew = новыйЛист.Cells(новыйЛист.Rows.count, "A").End(xlUp).Row
                            For j = 2 To lastRowNew ' Предполагаем, что первая строка - заголовок
                                If новыйЛист.Cells(j, 3).Value = wsСводная.Cells(i, 2).Value And _
                                   новыйЛист.Cells(j, 2).Value = ФИО Then
                                    заявкаУжеРаспределена = True
                                    Exit For
                                End If
                            Next j
                        End If

                        If Not заявкаУжеРаспределена Then
                            ' Если заявка не была распределена, добавляем ее на лист и в коллекцию
                            Dim lastRowNewList As Long
                            lastRowNewList = новыйЛист.Cells(новыйЛист.Rows.count, "A").End(xlUp).Row + 1
                            With новыйЛист
                                .Cells(lastRowNewList, 1).Value = DateValue(датаЗаявки) ' Дата заявки без времени
                                .Cells(lastRowNewList, 2).Value = ФИО ' ФИО сотрудника
                                .Cells(lastRowNewList, 3).Value = wsСводная.Cells(i, 2).Value ' Номер заявки
                                .Columns("A:C").AutoFit
                            End With

                            ' Увеличиваем счетчик распределенных заявок
                            распределеноЗаявок = распределеноЗаявок + 1
                            distributedRequests.Add wsСводная.Cells(i, 2).Value & "|" & ФИО

                            ' Переход к следующему сотруднику
                            currentСотрудник = currentСотрудник + 1
                            If currentСотрудник > countПрисутствующие Then currentСотрудник = 1 ' Сброс до первого сотрудника при необходимости
                        End If
                    End If
                End If
            End If
        End If
    Next i

    ' Обновляем индекс последнего обработанного сотрудника в ячейке (например AY2)
    wsСводная.Range("AY2").Value = currentСотрудник - 1 ' Сохраняем индекс последнего обработанного сотрудника

    MsgBox "Распределено заявок: " & распределеноЗаявок & vbNewLine & _
           "Присутствующих сотрудников: " & countПрисутствующие, vbInformation
End Sub


