' ==========================================
'  Excel × Outlook 自動化ツール
'  main.vba
'  ------------------------------------------
'  ・未読メールをExcelに抽出
'  ・Excelのデータを元にメールを自動送信
'  ・テンプレートファイル(template_mail.xlsx)に対応
' ==========================================

Option Explicit

' ==========================================
' 未読メールをExcelに抽出
' ==========================================
Sub ExtractUnreadEmails()
    Dim olApp As Outlook.Application
    Dim olNs As Outlook.Namespace
    Dim olFolder As Outlook.MAPIFolder
    Dim olMail As Outlook.MailItem
    Dim i As Long
    
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = New Outlook.Application
    On Error GoTo 0
    
    Set olNs = olApp.GetNamespace("MAPI")
    Set olFolder = olNs.GetDefaultFolder(olFolderInbox)
    
    ' 見出し行
    With ThisWorkbook.Sheets(1)
        .Cells(1, 1) = "受信日時"
        .Cells(1, 2) = "差出人"
        .Cells(1, 3) = "件名"
        .Cells(1, 4) = "本文"
        .Cells(1, 5) = "送信フラグ（1で送信）"
    End With
    
    i = 2
    For Each olMail In olFolder.Items
        If olMail.Class = olMail And olMail.UnRead Then
            With ThisWorkbook.Sheets(1)
                .Cells(i, 1) = olMail.ReceivedTime
                .Cells(i, 2) = olMail.SenderName
                .Cells(i, 3) = olMail.Subject
                .Cells(i, 4) = Left(olMail.Body, 200) '長すぎる本文は一部のみ抽出
            End With
            i = i + 1
        End If
    Next olMail
    
    MsgBox "未読メールの抽出が完了しました。", vbInformation
End Sub


' ==========================================
' Excel内容を元にメールを自動送信
' （template_mail.xlsx の本文テンプレート使用）
' ==========================================
Sub SendMailFromExcel()
    Dim olApp As Outlook.Application
    Dim olMail As Outlook.MailItem
    Dim ws As Worksheet
    Dim tempWb As Workbook
    Dim tempText As String
    Dim lastRow As Long, i As Long
    Dim templatePath As String
    
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = New Outlook.Application
    On Error GoTo 0
    
    Set ws = ThisWorkbook.Sheets(1)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' テンプレート読込
    templatePath = ThisWorkbook.Path & "\template_mail.xlsx"
    Set tempWb = Workbooks.Open(templatePath)
    tempText = Join(Application.Transpose(tempWb.Sheets(1).Range("A1:A15").Value), vbCrLf)
    tempWb.Close SaveChanges:=False
    
    ' メール送信
    For i = 2 To lastRow
        If ws.Cells(i, 5).Value = "1" Then
            Set olMail = olApp.CreateItem(olMailItem)
            With olMail
                .To = ws.Cells(i, 2).Value
                .Subject = ws.Cells(i, 3).Value
                .Body = Replace(tempText, "{body}", ws.Cells(i, 4).Value)
                .Send
            End With
            ws.Cells(i, 5).Value = "送信済"
        End If
    Next i
    
    MsgBox "メール送信が完了しました。", vbInformation
End Sub
