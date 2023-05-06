Sub ReplaceWordsFromExcel()
    ' 定义要替换的多组词
    Dim oldWords As Variant
    Dim newWords As Variant
    ' 获取用户输入的 Excel 文件路径
    Dim srcFilePath As String
    srcFilePath = InputBox("请输入替换词表文件路径")
    
    If srcFilePath = "" Then
        ' 如果用户没有输入任何路径，提醒用户并退出
        MsgBox "未输入路径"
        Exit Sub
    End If
    
    srcFilePath = Replace(srcFilePath, "/", "\")
    srcFilePath = Replace(srcFilePath, Chr(34), "")
    
    ' 打开 Excel 替换单词表
    Dim xlsApp As Object
    Dim xlsWorkbook As Object
    Dim xlsSheet As Object
    Set xlsApp = CreateObject("Excel.Application")
    Set xlsWorkbook = xlsApp.Workbooks.Open(srcFilePath)
    

    
    Set xlsSheet = xlsWorkbook.Sheets(1)
    oldWords = xlsSheet.Range("A1:A" & xlsSheet.Cells(xlsSheet.Rows.Count, "A").End(-4162).Row).Value
    newWords = xlsSheet.Range("B1:B" & xlsSheet.Cells(xlsSheet.Rows.Count, "B").End(-4162).Row).Value
    xlsWorkbook.Close False
    Set xlsSheet = Nothing
    Set xlsWorkbook = Nothing
    xlsApp.Quit
    Set xlsApp = Nothing
    ' 获取当前文档
    Dim doc As Document
    Set doc = ActiveDocument
    ' 替换所有标题
    ReplaceWordsInHeadersAndFooters doc, oldWords, newWords
    ' 替换所有表格
    Dim table As table
    For Each table In doc.Tables
        ReplaceWordsInText table.Range, oldWords, newWords
    Next table
    ' 替换正文
    ReplaceWordsInText doc.Content, oldWords, newWords
    ' 提示替换完成
    MsgBox "替换完成"
End Sub

' 在所有头/尾中替换词汇
Sub ReplaceWordsInHeadersAndFooters(doc As Document, oldWords As Variant, newWords As Variant)
    Dim storyRange As Range
    For Each storyRange In doc.StoryRanges
        If storyRange.StoryType = wdEvenPagesHeaderStory _
            Or storyRange.StoryType = wdFirstPageHeaderStory _
            Or storyRange.StoryType = wdPrimaryHeaderStory _
            Or storyRange.StoryType = wdEvenPagesFooterStory _
            Or storyRange.StoryType = wdFirstPageFooterStory _
            Or storyRange.StoryType = wdPrimaryFooterStory Then
            ReplaceWordsInText storyRange, oldWords, newWords
        End If
    Next storyRange
End Sub

' 在文本中替换词汇
Sub ReplaceWordsInText(text As Range, oldWords As Variant, newWords As Variant)
    Dim i As Long
    For i = LBound(oldWords) To UBound(oldWords)
        If oldWords(i, 1) <> "" Then
            ' 在文本中查找并替换所有匹配项
            With text.Find
                .ClearFormatting
                .MatchWholeWord = True
                .text = oldWords(i, 1)
                .Replacement.text = newWords(i, 1)
                .Execute Replace:=wdReplaceAll
            End With
        End If
    Next i
End Sub
