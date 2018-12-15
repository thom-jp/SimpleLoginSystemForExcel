Attribute VB_Name = "Module1"
Sub Login()
    Call ThisWorkbook.HideAll
    If LoginSheet.id = vbNullString Then
        MsgBox "ログインに失敗しました。"
        Exit Sub
    End If
    
    If LoginSheet.Password = vbNullString Then
        MsgBox "ログインに失敗しました。"
        Exit Sub
    End If
    
    Dim maxRow As Long: maxRow _
        = AccountsSheet.Range("A" & Rows.Count).End(xlUp).Row
    
    Dim i As Long
    For i = 4 To maxRow
        If UCase(AccountsSheet.Range("A" & i).Value) = UCase(LoginSheet.id) Then
            If AccountsSheet.Range("B" & i).Value = LoginSheet.Password Then
                'ID・パスワードがヒットしたら抜ける。
                Exit For
            End If
        End If
    Next
    
    'For文が最後まで回りきるとiが終値maxRowを
    '超えることを利用してログイン失敗を検知
    If i > maxRow Then
        MsgBox "ログインに失敗しました。"
        Exit Sub
    End If
    
    AccountsSheet.CurrentUser = LoginSheet.id
    Dim sheetName
    For Each sheetName In Split(AccountsSheet.Range("C" & i).Value, ":")
        Sheets(sheetName).Visible = xlSheetVisible
    Next
    LoginSheet.Activate
End Sub
