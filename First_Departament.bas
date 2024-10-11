Attribute VB_Name = "Module1"
Sub –аспределить«а€вки()
    Dim ws—водна€ As Worksheet, ws“абель As Worksheet
    Dim lastRow—водна€ As Long, lastRow“абель As Long, lastCol“абель As Long
    Dim выбранна€ƒата As Variant
    Dim присутствующие() As String
    Dim руководители() As String
    Dim countѕрисутствующие As Long
    Dim распределено«а€вок As Long
    Dim ‘»ќ As String
    Dim новыйЋист As Worksheet
    Dim безопасное»м€ As String
    Dim i As Long, j As Long, k As Long
    Dim current—отрудник As Long
    Dim за€вка”же–аспределена As Boolean

    ' «апрос даты у пользовател€
    выбранна€ƒата = InputBox("¬ведите дату дл€ распределени€ за€вок (в формате ƒƒ.ћћ.√√√√):", "¬ыбор даты")
    If Not IsDate(выбранна€ƒата) Then
        MsgBox "ќпераци€ отменена пользователем или введена некорректна€ дата.", vbInformation
        Exit Sub
    End If
    выбранна€ƒата = DateValue(выбранна€ƒата)

    ' ”становка листов
    Set ws—водна€ = ThisWorkbook.Sheets("—¬ќƒЌјя")
    Set ws“абель = ThisWorkbook.Sheets("“абель ѕлатежи")

    ' ќпределение последних строк и столбцов
    lastRow—водна€ = ws—водна€.Cells(ws—водна€.Rows.count, "A").End(xlUp).Row
    lastRow“абель = ws“абель.Cells(ws“абель.Rows.count, "B").End(xlUp).Row
    lastCol“абель = ws“абель.Cells(1, ws“абель.Columns.count).End(xlToLeft).Column

    ' —обрать список присутствующих сотрудников и их руководителей
    ReDim присутствующие(1 To lastRow“абель - 1)
    ReDim руководители(1 To lastRow“абель - 1)
    countѕрисутствующие = 0

    For j = 3 To lastCol“абель
        If IsDate(ws“абель.Cells(1, j).Value) And DateValue(ws“абель.Cells(1, j).Value) = выбранна€ƒата Then
            For k = 2 To lastRow“абель
                If ws“абель.Cells(k, j).Value = "ƒа" Then
                    countѕрисутствующие = countѕрисутствующие + 1
                    присутствующие(countѕрисутствующие) = ws“абель.Cells(k, "B").Value ' —охран€ем ‘»ќ сотрудника
                    руководители(countѕрисутствующие) = ws“абель.Cells(k, "A").Value ' —охран€ем им€ руководител€

                    безопасное»м€ = Application.WorksheetFunction.Substitute(руководители(countѕрисутствующие), "/", " ")
                    On Error Resume Next
                    Set новыйЋист = ThisWorkbook.Sheets(безопасное»м€)
                    On Error GoTo 0

                    ' —оздаем новый лист дл€ руководител€, если он не существует
                    If новыйЋист Is Nothing Then
                        Set новыйЋист = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
                        новыйЋист.Name = безопасное»м€
                    End If

                    Set новыйЋист = Nothing
                End If
            Next k
            Exit For
        End If
    Next j

    If countѕрисутствующие = 0 Then
        MsgBox "Ќа выбранную дату нет присутствующих сотрудников.", vbExclamation
        Exit Sub
    End If

    If IsEmpty(ws—водна€.Range("AY1").Value) Then
        current—отрудник = 1
    Else
        current—отрудник = ws—водна€.Range("AY1").Value
        If current—отрудник < 1 Then current—отрудник = 1
        If current—отрудник > countѕрисутствующие Then current—отрудник = 1
    End If

    For i = 4 To lastRow—водна€
        Dim дата«а€вки As Variant
        дата«а€вки = ws—водна€.Cells(i, 1).Value

        If IsDate(дата«а€вки) Then
            If DateValue(дата«а€вки) = выбранна€ƒата Then
                If ws—водна€.Cells(i, "Y").Value = "" Then
                    ' ѕолучаем им€ руководител€ из массива
                    безопасное»м€ = Application.WorksheetFunction.Substitute(руководители(current—отрудник), "/", " ")
                    On Error Resume Next
                    Set новыйЋист = ThisWorkbook.Sheets(безопасное»м€)
                    On Error GoTo 0

                    ' ѕровер€ем, существует ли лист руководител€ и не была ли за€вка уже распределена
                    If Not новыйЋист Is Nothing And Not IsEmpty(присутствующие(current—отрудник)) Then
                        за€вка”же–аспределена = False
                        Dim lastRowNew As Long
                        lastRowNew = новыйЋист.Cells(новыйЋист.Rows.count, "A").End(xlUp).Row
                        
                        For j = 2 To lastRowNew ' ѕредполагаем, что перва€ строка - заголовок
                            If новыйЋист.Cells(j, 3).Value = ws—водна€.Cells(i, 2).Value Then
                                за€вка”же–аспределена = True
                                Exit For
                            End If
                        Next j

                        If Not за€вка”же–аспределена Then
                            ‘»ќ = присутствующие(current—отрудник)
                            ws—водна€.Cells(i, "X").Value = руководители(current—отрудник) ' –уководитель
                            ws—водна€.Cells(i, "Y").Value = ‘»ќ ' —отрудник

                            распределено«а€вок = распределено«а€вок + 1

                            With новыйЋист
                                Dim lastRowNewList As Long
                                lastRowNewList = .Cells(.Rows.count, "A").End(xlUp).Row + 1
                                .Cells(lastRowNewList, 1).Value = DateValue(дата«а€вки)
                                .Cells(lastRowNewList, 2).Value = ‘»ќ
                                .Cells(lastRowNewList, 3).Value = ws—водна€.Cells(i, 2).Value ' Ќомер за€вки
                                .Columns("A:C").AutoFit
                            End With
                        End If
                    End If
                End If
            End If
        End If

        ' ѕереключаемс€ на следующего сотрудника вне зависимости от результата
        current—отрудник = current—отрудник + 1
        If current—отрудник > countѕрисутствующие Then current—отрудник = 1
    Next i

    ws—водна€.Range("AY1").Value = current—отрудник

    MsgBox "–аспределено за€вок: " & распределено«а€вок & vbNewLine & _
           "ѕрисутствующих сотрудников: " & countѕрисутствующие, vbInformation
End Sub

