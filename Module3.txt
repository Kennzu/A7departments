Attribute VB_Name = "Module1"
Sub РаспределитьЗаявки()
    Dim wsСводная As Worksheet, wsТабель As Worksheet
    Dim lastRowСводная As Long, lastRowТабель As Long, lastColТабель As Long
    Dim выбраннаяДата As Variant
    Dim присутствующие() As String
    Dim руководители() As String
    Dim countПрисутствующие As Long
    Dim распределеноЗаявок As Long
    Dim ФИО As String
    Dim новыйЛист As Worksheet
    Dim безопасноеИмя As String
    Dim i As Long, j As Long, k As Long
    Dim currentСотрудник As Long
    Dim заявкаУжеРаспределена As Boolean

    ' Запрос даты у пользователя
    выбраннаяДата = InputBox("Введите дату для распределения заявок (в формате ДД.ММ.ГГГГ):", "Выбор даты")
    If Not IsDate(выбраннаяДата) Then
        MsgBox "Операция отменена пользователем или введена некорректная дата.", vbInformation
        Exit Sub
    End If
    выбраннаяДата = DateValue(выбраннаяДата)

    ' Установка листов
    Set wsСводная = ThisWorkbook.Sheets("СВОДНАЯ")
    Set wsТабель = ThisWorkbook.Sheets("Табель Платежи")

    ' Определение последних строк и столбцов
    lastRowСводная = wsСводная.Cells(wsСводная.Rows.count, "A").End(xlUp).Row
    lastRowТабель = wsТабель.Cells(wsТабель.Rows.count, "B").End(xlUp).Row
    lastColТабель = wsТабель.Cells(1, wsТабель.Columns.count).End(xlToLeft).Column

    ' Собрать список присутствующих сотрудников и их руководителей
    ReDim присутствующие(1 To lastRowТабель - 1)
    ReDim руководители(1 To lastRowТабель - 1)
    countПрисутствующие = 0

    For j = 3 To lastColТабель
        If IsDate(wsТабель.Cells(1, j).Value) And DateValue(wsТабель.Cells(1, j).Value) = выбраннаяДата Then
            For k = 2 To lastRowТабель
                If wsТабель.Cells(k, j).Value = "Да" Then
                    countПрисутствующие = countПрисутствующие + 1
                    присутствующие(countПрисутствующие) = wsТабель.Cells(k, "B").Value ' Сохраняем ФИО сотрудника
                    руководители(countПрисутствующие) = wsТабель.Cells(k, "A").Value ' Сохраняем имя руководителя

                    безопасноеИмя = Application.WorksheetFunction.Substitute(руководители(countПрисутствующие), "/", " ")
                    On Error Resume Next
                    Set новыйЛист = ThisWorkbook.Sheets(безопасноеИмя)
                    On Error GoTo 0

                    ' Создаем новый лист для руководителя, если он не существует
                    If новыйЛист Is Nothing Then
                        Set новыйЛист = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
                        новыйЛист.Name = безопасноеИмя
                    End If

                    Set новыйЛист = Nothing
                End If
            Next k
            Exit For
        End If
    Next j

    If countПрисутствующие = 0 Then
        MsgBox "На выбранную дату нет присутствующих сотрудников.", vbExclamation
        Exit Sub
    End If

    If IsEmpty(wsСводная.Range("AY1").Value) Then
        currentСотрудник = 1
    Else
        currentСотрудник = wsСводная.Range("AY1").Value
        If currentСотрудник < 1 Then currentСотрудник = 1
        If currentСотрудник > countПрисутствующие Then currentСотрудник = 1
    End If

    For i = 4 To lastRowСводная
        Dim датаЗаявки As Variant
        датаЗаявки = wsСводная.Cells(i, 1).Value

        If IsDate(датаЗаявки) Then
            If DateValue(датаЗаявки) = выбраннаяДата Then
                If wsСводная.Cells(i, "Y").Value = "" Then
                    ' Получаем имя руководителя из массива
                    безопасноеИмя = Application.WorksheetFunction.Substitute(руководители(currentСотрудник), "/", " ")
                    On Error Resume Next
                    Set новыйЛист = ThisWorkbook.Sheets(безопасноеИмя)
                    On Error GoTo 0

                    ' Проверяем, существует ли лист руководителя и не была ли заявка уже распределена
                    If Not новыйЛист Is Nothing And Not IsEmpty(присутствующие(currentСотрудник)) Then
                        заявкаУжеРаспределена = False
                        Dim lastRowNew As Long
                        lastRowNew = новыйЛист.Cells(новыйЛист.Rows.count, "A").End(xlUp).Row
                        
                        For j = 2 To lastRowNew ' Предполагаем, что первая строка - заголовок
                            If новыйЛист.Cells(j, 3).Value = wsСводная.Cells(i, 2).Value Then
                                заявкаУжеРаспределена = True
                                Exit For
                            End If
                        Next j

                        If Not заявкаУжеРаспределена Then
                            ФИО = присутствующие(currentСотрудник)
                            wsСводная.Cells(i, "X").Value = руководители(currentСотрудник) ' Руководитель
                            wsСводная.Cells(i, "Y").Value = ФИО ' Сотрудник

                            распределеноЗаявок = распределеноЗаявок + 1

                            With новыйЛист
                                Dim lastRowNewList As Long
                                lastRowNewList = .Cells(.Rows.count, "A").End(xlUp).Row + 1
                                .Cells(lastRowNewList, 1).Value = DateValue(датаЗаявки)
                                .Cells(lastRowNewList, 2).Value = ФИО
                                .Cells(lastRowNewList, 3).Value = wsСводная.Cells(i, 2).Value ' Номер заявки
                                .Columns("A:C").AutoFit
                            End With
                        End If
                    End If
                End If
            End If
        End If

        ' Переключаемся на следующего сотрудника вне зависимости от результата
        currentСотрудник = currentСотрудник + 1
        If currentСотрудник > countПрисутствующие Then currentСотрудник = 1
    Next i

    wsСводная.Range("AY1").Value = currentСотрудник

    MsgBox "Распределено заявок: " & распределеноЗаявок & vbNewLine & _
           "Присутствующих сотрудников: " & countПрисутствующие, vbInformation
End Sub

