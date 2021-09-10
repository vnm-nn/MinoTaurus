VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "ФОРМА ЗАПОЛНЕНИЯ v2.0"
   ClientHeight    =   7150
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   13350
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public EnebleEvents As Boolean


Private Sub cmdSearchColumn_Change()

    If Me.EnebleEvents = False Then Exit Sub ' была ошибка, но пофикшено
    If Me.cmbSearchColumn.Value = "All" Then
    
        Call Reset
        
    Else
        
        Me.txtSearch.Value = ""
        Me.txtSearch.Enabled = True
        Me.cmdSearch.Enabled = True
             
    End If
    

End Sub

Private Sub cmdDelet_Click()

    Dim iRow As Long
    
    If Selected_List = 0 Then
    
        MsgBox "Строка не выбрана!", vbOKOnly + vbInformation, "Удаление записи"
        
        Exit Sub
        
    End If
    
    Dim i As VbMsgBoxResult
    
    i = MsgBox("Вы уверены что хотите удалить запись?", vbYesNo + vbQuestion, "Удаление записи")
    
    If i = vbNo Then Exit Sub
    
    iRow = Application.WorksheetFunction.Match(Me.lstDataBase.List(Me.lstDataBase.ListIndex, 0), _
    ThisWorkbook.Sheets("DataBase").Range("A:A"), 0)
    
    ThisWorkbook.Sheets("DataBase").Rows(iRow).Delete

    Call Reset
    
    MsgBox "Выбранная строка удалена.", vbOKOnly + vbInformation, "Удаление записи"
    
End Sub

Private Sub cmdEdit_Click()

    If Selected_List = 0 Then
    
        MsgBox "Строка не выбрана!", vbOKOnly + vbInformation, "Редактировать"
        
        Exit Sub
    
    End If
    
    'Участок для обновления значения
    
    Me.txtRowNumber.Value = Application.WorksheetFunction.Match(Me.lstDataBase.List(Me.lstDataBase.ListIndex, 0), _
    ThisWorkbook.Sheets("DataBase").Range("A:A"), 0)
    
    Me.ComboBox1.Value = Me.lstDataBase.List(Me.lstDataBase.ListIndex, 1)
    Me.TextBox1.Value = Me.lstDataBase.List(Me.lstDataBase.ListIndex, 2)
    Me.ComboBox2.Value = Me.lstDataBase.List(Me.lstDataBase.ListIndex, 3)
    Me.ComboBox3.Value = Me.lstDataBase.List(Me.lstDataBase.ListIndex, 4)
    Me.TextBox2.Value = Me.lstDataBase.List(Me.lstDataBase.ListIndex, 5)
    Me.TextBox3.Value = Me.lstDataBase.List(Me.lstDataBase.ListIndex, 6)
    Me.TextBox4.Value = Me.lstDataBase.List(Me.lstDataBase.ListIndex, 7)
    
    MsgBox "Пожалуйста измените поля и нажмите кнопку 'Сохранить' для обновления записи", vbOKOnly + vbInformation, "Редактировать"
    

End Sub

Private Sub cmdReset_Click()

    Dim msgValue As VbMsgBoxResult
    
    msgValue = MsgBox("ВНИМАНИЕ Вы хотите стереть информацию в форме?", vbYesNo + vbInformation, "Удалено!")
    
    If msgValue = vbNo Then Exit Sub

    Call Reset
    
End Sub

Private Sub cmdSave_Click()

    Dim msgValue As VbMsgBoxResult
    
    msgValue = MsgBox("Вы хотите сохратить информацтю в форме?", vbYesNo + vbInformation, "Сохранено!")
    
    If msgValue = vbNo Then Exit Sub
    
    Call Submit
    Call Reset
     
End Sub

Private Sub txtRowNunber_Change()

End Sub

Private Sub ComboBox4_Change()

End Sub

Private Sub cmdSearch_Click()

    If Me.txtSearch.Value = "" Then
    
        MsgBox "Введите значение для поиска", vbOKOnly + vbInformation, "Поиск"
        
        Exit Sub
            
    End If

    Call SearchData

End Sub

Private Sub UserForm_Initialize()

    ' Aplication.Visible = False
    ' UserForm.Show
        
    Call Reset

End Sub

' Private Sub UserForm_Terminate()

'     Aplication.Visible = True

' End Sub

'-----------------------------------------
' для скрола мышом

Private Sub ComboBox1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    HookScroll Me.ComboBox1

End Sub

Private Sub ComboBox2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    HookScroll Me.ComboBox2
    
End Sub

Private Sub ComboBox3_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    HookScroll Me.ComboBox3
    
End Sub

Private Sub cmbSearchColumn_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    HookScroll Me.cmbSearchColumn
    
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    UnHookScroll
    
End Sub

' для скрола
'---------------------------------------



