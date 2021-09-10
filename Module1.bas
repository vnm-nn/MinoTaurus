Attribute VB_Name = "Module1"
Sub Reset()

    Dim iRow As Long
    
    iRow = [Counta(DataBase!A:A)] '������������� ���������� ��������
    
    With UserForm
    
        .TextBox1.Value = ".��"
        .TextBox2.Value = ".��"
        .TextBox3.Value = ".��"
        .TextBox4.Value = ".��"
            
        .ComboBox1.Clear
        
        .ComboBox1.AddItem "1�"
        .ComboBox1.AddItem "2�"
        .ComboBox1.AddItem "3�"
        .ComboBox1.AddItem "4�"
        .ComboBox1.AddItem "5�"
        .ComboBox1.AddItem "6�"
        .ComboBox1.AddItem "7�"
        .ComboBox1.AddItem "8�"
        .ComboBox1.AddItem "9�"
        .ComboBox1.AddItem "10�"
        .ComboBox1.AddItem "11�"
        .ComboBox1.AddItem "12�"
        .ComboBox1.AddItem "13�"
        .ComboBox1.AddItem "14�"
        .ComboBox1.AddItem "15�"
        .ComboBox1.AddItem "16�"
        .ComboBox1.AddItem "17�"
        
        
        .ComboBox2.Clear
        
        .ComboBox2.AddItem "1q"
        .ComboBox2.AddItem "2q"
        .ComboBox2.AddItem "3q"
        .ComboBox2.AddItem "4q"
        .ComboBox2.AddItem "5q"
        .ComboBox2.AddItem "6q"
        .ComboBox2.AddItem "7q"
        .ComboBox2.AddItem "8q"
        .ComboBox2.AddItem "9q"
        .ComboBox2.AddItem "10q"
        .ComboBox2.AddItem "11q"
        .ComboBox2.AddItem "12q"
        .ComboBox2.AddItem "13q"
        .ComboBox2.AddItem "14q"
        .ComboBox2.AddItem "15q"
        .ComboBox2.AddItem "16q"
        .ComboBox2.AddItem "17q"
        .ComboBox2.AddItem "18q"
        .ComboBox2.AddItem "19q"
        .ComboBox2.AddItem "20q"
        .ComboBox2.AddItem "21q"
        .ComboBox2.AddItem "22q"
        .ComboBox2.AddItem "23q"
        
        
        .ComboBox3.Clear
        
        .ComboBox3.AddItem "1999"
        .ComboBox3.AddItem "2000"
        .ComboBox3.AddItem "2001"
        .ComboBox3.AddItem "2002"
        .ComboBox3.AddItem "2003"
        .ComboBox3.AddItem "2004"
        .ComboBox3.AddItem "2005"
        .ComboBox3.AddItem "2006"
        .ComboBox3.AddItem "2007"
        .ComboBox3.AddItem "2008"
        .ComboBox3.AddItem "2009"
        .ComboBox3.AddItem "2010"
        .ComboBox3.AddItem "2011"
        .ComboBox3.AddItem "2012"
        .ComboBox3.AddItem "2013"
        .ComboBox3.AddItem "2014"
        .ComboBox3.AddItem "2015"
        .ComboBox3.AddItem "2016"
        .ComboBox3.AddItem "2017"
        .ComboBox3.AddItem "2018"
        .ComboBox3.AddItem "2019"
        .ComboBox3.AddItem "2020"
        .ComboBox3.AddItem "2021"
        
        .txtRowNumber.Value = "" ' ������� ��� ������ ��������� ��� (�������)
        
        ' ��� ����� ����� ��� ���� ��� �� ������� �����
        Call Add_SearchColumn
        ThisWorkbook.Sheets("DataBase").AutoFilterMode = False
        ThisWorkbook.Sheets("SearchData").AutoFilterMode = False
        ThisWorkbook.Sheets("SearchData").Cells.Clear
        
        '---------------------------------------------
                
        .lstDataBase.ColumnCount = 9
        .lstDataBase.ColumnHeads = True
        
        .lstDataBase.ColumnWidths = "40, 60, 70, 50, 60, 60, 70, 70, 60"
        
        If iRow > 1 Then
            .lstDataBase.RowSource = "DataBase!A2:I" & iRow
        Else
            .lstDataBase.RowSource = "DataBase!A2:I2"
        End If
    
    End With
    
End Sub

Sub Submit()
    
    Dim sh As Worksheet
    Dim iRow As Long
    
    Set sh = ThisWorkbook.Sheets("DataBase")
    
    If UserForm.txtRowNumber.Value = "" Then
    
        iRow = [Counta(DataBase!A:A)] + 1
    
    Else
    
        iRow = UserForm.txtRowNumber.Value
        
    End If
    
    With sh ' ������� ������� ������� ����� ������������ � �����
    
        .Cells(iRow, 1) = iRow - 1
        
        .Cells(iRow, 2) = UserForm.ComboBox1.Value
        
        .Cells(iRow, 3) = UserForm.TextBox1.Value
        
        .Cells(iRow, 4) = UserForm.ComboBox2.Value
        
        .Cells(iRow, 5) = UserForm.ComboBox3.Value
        
        .Cells(iRow, 6) = UserForm.TextBox2.Value
        
        .Cells(iRow, 7) = UserForm.TextBox3.Value
        
        .Cells(iRow, 8) = UserForm.TextBox4.Value
        
        .Cells(iRow, 9) = [Text(Now(), "dd-mm-yy hh:mm")]
    
    End With
            
End Sub

Sub Show_Form()

    UserForm.Show

End Sub

Function Selected_List() As Long

    Dim i As Long
    
    Selected_List = 0
    
    For i = 0 To UserForm.lstDataBase.ListCount - 1
    
        If UserForm.lstDataBase.Selected(i) = True Then
        
            Selected_List = i + 1
            Exit For
        End If
            
    Next i
        
End Function

Sub Add_SearchColumn()

    UserForm.EnebleEvents = False
    
    With UserForm.cmbSearchColumn
    
        .Clear
        
        .AddItem "All"
        
        .AddItem "��� �������"
        .AddItem "�������� �� ������"
        .AddItem "������������"
        .AddItem "��� �������"
        .AddItem "���-��"
        .AddItem "������� � ������� ����"
        .AddItem "����������� � �������� � ������� ����"
        .AddItem "���� ����������"
        
        .Value = "All"
        
    End With
            
    UserForm.EnebleEvents = True
        
    UserForm.txtSearch.Value = ""
    UserForm.txtSearch.Enabled = False
    UserForm.cmdSearch.Enabled = False
    

End Sub


Sub SearchData()

    Application.ScreenUpdating = False
    
    Dim shDataBase As Worksheet ' ���� � ��
    Dim shSearchData As Worksheet ' ���� ��� ������
    
    Dim iColumn As Integer
    Dim iDataBaseRow As Long
    Dim iSearchRow As Long
    
    Dim sColumn As String
    Dim sValue As String
    
    Set shDataBase = ThisWorkbook.Sheets("DataBase")
    Set shSearchData = ThisWorkbook.Sheets("SearchData")
     
    iDataBaseRow = ThisWorkbook.Sheets("DataBase").Range("A" & Aplication.Rows.Count).End(xlUp).Row
    
    sColumn = UserForm.cmbSearchColumn.Value
    sValue = UserForm.txtSearch.Value
    
    iColumn = Application.WorksheetFunction.Match(sColumn, shDataBase.Range("A1:I1"), 0)
       
    ' ����� �������
    
    If shDataBase.FilterMode = True Then
        
        shDataBase.AutoFilterMode = False
        
    End If
        
    ' �� �������
    
    If UserForm.cmbSearchColumn.Value = "All" Then ' �������� �� ������ ������
    
        shDataBase.Range("A1:I" & DataBaseRow).AutoFilter Field:=iColumn, Criteria1:=sValue
        
    Else
        
        shDataBase.Range("A1:I" & DataBaseRow).AutoFilter Field:=iColumn, Criteria1:="*" & sValue & "*"
                
    End If
    
    If Application.WorksheetFunction.Subtotal(3, shDataBase.Range("C:C")) >= 2 Then
    
        shSearchData.Cells.Clear
        
        shDataBase.AutoFilter.Range.Copy shSearchData.Range("A1")
        
        Application.CutCopyMode = False
        
        iSearchRow = shSearchData.Range("A" & Application.Rows.Count).End(xlUp).Row
        
        UserForm.lstDataBase.ColumnCount = 9
        
        UserForm.lstDataBase.ColumnWidths = "40, 60, 70, 50, 60, 60, 70, 70, 60"
        
        If iSearchRow > 1 Then
        
            UserForm.lstDataBase.RowSource = "SearchData!A2:I" & iSearchRow
            
            MsgBox "������ �������."
            
        End If
        
    
    Else
    
        MsgBox "������ �� �������."
    
    End If
    
    shDataBase.AutoFilterMode = False
    
    Application.ScreenUpdating = True
    
End Sub





























































