Attribute VB_Name = "Module1"
Sub Login()
    Call ThisWorkbook.HideAll
    If LoginSheet.id = vbNullString Then
        MsgBox "���O�C���Ɏ��s���܂����B"
        Exit Sub
    End If
    
    If LoginSheet.Password = vbNullString Then
        MsgBox "���O�C���Ɏ��s���܂����B"
        Exit Sub
    End If
    
    Dim maxRow As Long: maxRow _
        = AccountsSheet.Range("A" & Rows.Count).End(xlUp).Row
    
    Dim i As Long
    For i = 4 To maxRow
        If UCase(AccountsSheet.Range("A" & i).Value) = UCase(LoginSheet.id) Then
            If AccountsSheet.Range("B" & i).Value = LoginSheet.Password Then
                'ID�E�p�X���[�h���q�b�g�����甲����B
                Exit For
            End If
        End If
    Next
    
    'For�����Ō�܂ŉ�肫���i���I�lmaxRow��
    '�����邱�Ƃ𗘗p���ă��O�C�����s�����m
    If i > maxRow Then
        MsgBox "���O�C���Ɏ��s���܂����B"
        Exit Sub
    End If
    
    AccountsSheet.CurrentUser = LoginSheet.id
    Dim sheetName
    For Each sheetName In Split(AccountsSheet.Range("C" & i).Value, ":")
        Sheets(sheetName).Visible = xlSheetVisible
    Next
    LoginSheet.Activate
End Sub
