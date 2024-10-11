Attribute VB_Name = "Module2"
Sub ��������������������������()
    Dim ws������� As Worksheet, ws������ As Worksheet
    Dim lastRow������� As Long, lastRow������ As Long, lastCol������ As Long
    Dim ������������� As Variant
    Dim ��������������() As String
    Dim count�������������� As Long
    Dim ������������������ As Long
    Dim ��� As String
    Dim ��������� As Worksheet
    Dim ������������� As String
    Dim i As Long, j As Long, k As Long
    Dim lastProcessedEmployeeIndex As Long
    
    ' ��������� ������
    Set ws������� = ThisWorkbook.Sheets("�������")
    Set ws������ = ThisWorkbook.Sheets("������ ���������������")

    ' ��������� ������ ���������� ������������� ���������� �� ������ (��������, AY2)
    lastProcessedEmployeeIndex = ws�������.Range("AY2").Value

    ' ������ ���� � ������������
    ������������� = InputBox("������� ���� ��� ������������� ������ (� ������� ��.��.����):", "����� ����")
    If Not IsDate(�������������) Then
        MsgBox "�������� �������� ������������� ��� ������� ������������ ����.", vbInformation
        Exit Sub
    End If
    ������������� = DateValue(�������������)

    ' ����������� ��������� ����� � ��������
    lastRow������� = ws�������.Cells(ws�������.Rows.count, "A").End(xlUp).Row
    lastRow������ = ws������.Cells(ws������.Rows.count, "B").End(xlUp).Row
    lastCol������ = ws������.Cells(1, ws������.Columns.count).End(xlToLeft).Column

    ' ������� ������ �������������� �����������
    ReDim ��������������(1 To lastRow������ - 1)
    count�������������� = 0

    For j = 3 To lastCol������ ' �������� � C1
        If IsDate(ws������.Cells(1, j).Value) And DateValue(ws������.Cells(1, j).Value) = ������������� Then
            ' ������� ������ �������������� �����������
            For k = 2 To lastRow������
                If ws������.Cells(k, j).Value = "��" Then
                    count�������������� = count�������������� + 1
                    ��������������(count��������������) = ws������.Cells(k, "B").Value ' ��� �� ������� B

                    ' ������� ���������� ��� ��� ����� (�� ������������)
                    ������������� = Application.WorksheetFunction.Substitute(ws������.Cells(k, "A").Value, "/", " ")

                    ' ���������, ���������� �� ����, ���� ���, ������� ���
                    On Error Resume Next
                    Set ��������� = ThisWorkbook.Sheets(�������������)
                    On Error GoTo 0

                    ' ���� ���� �� ����������, ������� ���
                    If ��������� Is Nothing Then
                        Set ��������� = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
                        ���������.Name = �������������
                    End If

                    ' ����� ������ �� ����� ���� ��� ��������� ��������
                    Set ��������� = Nothing
                End If
            Next k
            Exit For ' ������� �� ����� ����� ���������� ����
        End If
    Next j

    If count�������������� = 0 Then
        MsgBox "�� ��������� ���� ��� �������������� �����������.", vbExclamation
        Exit Sub
    End If

    ' ������� ��������� ��� �������� ��� �������������� ������� ������ � ��������������� �����������
    Dim distributedRequests As Collection
    Set distributedRequests = New Collection

    ' ��������� ��������� ��� ��������������� �������� ������
    For i = 4 To lastRow�������
        If ws�������.Cells(i, "T").Value <> "" Then ' ���� ������ T �� ������
            distributedRequests.Add ws�������.Cells(i, 2).Value & "|" & ws�������.Cells(i, "T").Value ' ����� ������ � ��� ����������
        End If
    Next i

    ' ������������ ������ �� ��� ����� ������� � ���������� ������������� ���������� + 1
    Dim current��������� As Long
    current��������� = lastProcessedEmployeeIndex + 1

    For i = 4 To lastRow������� ' �������� � A4 ��� ������
        Dim ���������� As Variant
        ���������� = ws�������.Cells(i, 1).Value ' ���� ������ �� ������� A
        
        If IsDate(����������) Then
            ' ���������� ������ ���� ��� ����� ������� � ��������� ����� ������
            If DateValue(����������) = ������������� Then
                
                ' �������� �� ������� ������ ������ � ���������
                Dim requestExists As Boolean
                requestExists = False

                For Each Item In distributedRequests
                    If InStr(Item, ws�������.Cells(i, 2).Value) > 0 Then
                        requestExists = True
                        Exit For
                    End If
                Next Item

                If Not requestExists Then
                    ' ���������, ���� �� ��� ������ � ������� T ��� ���� ������
                    If ws�������.Cells(i, "T").Value = "" Then ' ���� ������ ������, ������ �� ������������
                        ' �������� ��� ����������
                        ��� = ��������������(current���������)
                        ' ���������� ������ � ��������������� ������
                        ws�������.Cells(i, "S").Value = ws������.Cells(Application.Match(���, ws������.Range("B:B"), 0), "A").Value ' ������������
                        ws�������.Cells(i, "T").Value = ��� ' ���������

                        ' �������� �� ������� ������ �� ����� ������������
                        ������������� = Application.WorksheetFunction.Substitute(ws������.Cells(Application.Match(���, ws������.Range("B:B"), 0), "A").Value, "/", " ")
                        
                        On Error Resume Next
                        Set ��������� = ThisWorkbook.Sheets(�������������)
                        On Error GoTo 0

                        Dim ��������������������� As Boolean
                        ��������������������� = False

                        ' ���������, ���������� �� ������ �� ����� ������������
                        If Not ��������� Is Nothing Then
                            Dim lastRowNew As Long
                            lastRowNew = ���������.Cells(���������.Rows.count, "A").End(xlUp).Row
                            For j = 2 To lastRowNew ' ������������, ��� ������ ������ - ���������
                                If ���������.Cells(j, 3).Value = ws�������.Cells(i, 2).Value And _
                                   ���������.Cells(j, 2).Value = ��� Then
                                    ��������������������� = True
                                    Exit For
                                End If
                            Next j
                        End If

                        If Not ��������������������� Then
                            ' ���� ������ �� ���� ������������, ��������� �� �� ���� � � ���������
                            Dim lastRowNewList As Long
                            lastRowNewList = ���������.Cells(���������.Rows.count, "A").End(xlUp).Row + 1
                            With ���������
                                .Cells(lastRowNewList, 1).Value = DateValue(����������) ' ���� ������ ��� �������
                                .Cells(lastRowNewList, 2).Value = ��� ' ��� ����������
                                .Cells(lastRowNewList, 3).Value = ws�������.Cells(i, 2).Value ' ����� ������
                                .Columns("A:C").AutoFit
                            End With

                            ' ����������� ������� �������������� ������
                            ������������������ = ������������������ + 1
                            distributedRequests.Add ws�������.Cells(i, 2).Value & "|" & ���

                            ' ������� � ���������� ����������
                            current��������� = current��������� + 1
                            If current��������� > count�������������� Then current��������� = 1 ' ����� �� ������� ���������� ��� �������������
                        End If
                    End If
                End If
            End If
        End If
    Next i

    ' ��������� ������ ���������� ������������� ���������� � ������ (�������� AY2)
    ws�������.Range("AY2").Value = current��������� - 1 ' ��������� ������ ���������� ������������� ����������

    MsgBox "������������ ������: " & ������������������ & vbNewLine & _
           "�������������� �����������: " & count��������������, vbInformation
End Sub


