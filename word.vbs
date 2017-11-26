' 查找最近的段落编号
Function findNumber()
    Dim prev As Object

    While True
        Set prev = ActiveDocument.ActiveWindow.Selection.Previous(wdParagraph)
        If InStr(prev.ListFormat.ListString, ".") Then
            findNumber = prev.ListFormat.ListString
            Set prev = Nothing
            Exit Function
        Else
            prev.Select
        End If
    Wend
    
End Function


' 自动统一编号项
Sub 高级需求统一编号()
  
     Dim snItems
     Dim snSuffix As String
     
     For Each aTable In ActiveDocument.Tables
        If aTable.Rows.Count >= 2 Then
             If InStr(aTable.Cell(1, 1).Range.Text, "编号") > 0 Then
                aTable.Range.Select
                sn = findNumber()
                snItems = Split(sn, ".", -1, 1)
                 
                snSuffix = "HLR_"
                For i = 0 To UBound(snItems)
                     snSuffix = snSuffix + snItems(i) + "_"
                Next
                 
                For i = 1 To aTable.Rows.Count - 1
                    nn = Str(i)
                    aTable.Cell(i + 1, 1).Range.Text = snSuffix + Right(nn, Len(nn) - 1)
                 Next
             End If
         End If
     Next

End Sub

' 自动统一编号项
Sub 详细需求统一编号()
  
     Dim snItems
     Dim snSuffix As String
     
     For Each aTable In ActiveDocument.Tables
        If aTable.Rows.Count >= 2 And aTable.Columns.Count = 2 Then
             If InStr(aTable.Cell(1, 1).Range.Text, "编号") > 0 Then
                aTable.Range.Select
                sn = findNumber()
                snItems = Split(sn, ".", -1, 1)
                 
                snSuffix = "SRS_"
                For i = 0 To UBound(snItems)
                     snSuffix = snSuffix + snItems(i) + "_"
                Next
                
                nn = Str(1)
                aTable.Cell(1, 1 + 1).Range.Text = Left(snSuffix, Len(snSuffix) - 1)
                
             End If
         End If
     Next
    
End Sub

Function isTableName(aRange As Word.Range)

    isTableName = False
    
    If Left(aRange.Style, 2) = "标题" And Left(aRange.Text, 3) = "TB_" Then
        isTableName = True
    End If
    
End Function

Function convertTableToCSV(aTable As Word.Table)

    convertTableToCSV = ""

    If aTable.Rows.Count < 2 Then
        Exit Function
    End If
    
    If InStr(aTable.Cell(1, 1).Range.Text, "字段名") < 1 Then
        Exit Function
    End If

    

    For i = 1 To aTable.Rows.Count
        For j = 1 To aTable.Columns.Count
            convertTableToCSV = convertTableToCSV + clearString(aTable.Cell(i, j).Range.Text) + ","
        Next
        convertTableToCSV = convertTableToCSV + vbCrLf
    Next
    
End Function

Function clearString(aString As String)

    Dim re As Object
    Set re = CreateObject("vbscript.regexp")
    re.Pattern = "[\s,]"
    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = True
    
    clearString = re.Replace(aString, "")

End Function

Sub 复制数据库工具库可识别的CSV到粘贴板()

    aCSVFile = ""
    
    Dim aRange As Word.Range
    Set aRange = ActiveDocument.Sections(3).Range
    
    Dim paragraphsCount As Integer
    paragraphsCount = aRange.Paragraphs.Count

    For i = 1 To paragraphsCount Step 1
    
        Dim aParagraph As Word.Paragraph
        Set aParagraph = aRange.Paragraphs(i)
        
        Dim thisRange As Word.Range
        Set thisRange = aParagraph.Range
       
        If isTableName(thisRange) Then

            TableName = ""
            tableCaption = ""
            tableDeclare = ""
            
            TableName = clearString(thisRange.Text)
            
            If i >= paragraphsCount Then
                Exit For
            End If
            

            For j = i + 1 To paragraphsCount Step 1
            
                Dim nextRange As Word.Range
                Set nextRange = aRange.Paragraphs(j).Range
                
                If isTableName(nextRange) Then
                    
                    ' 后退一个索引，因为当前索引已经是下一个表了，由下个循环处理
                    i = j - 1

                    Exit For
                    
                ElseIf nextRange.Information(wdWithInTable) Then
                        
                    Dim thisTable As Word.Table
                    Set thisTable = nextRange.Tables(1)
                    
                    tableDeclare = convertTableToCSV(thisTable)
                    
                    cellCount = thisTable.Rows.Count * thisTable.Columns.Count
                    
                    
                    ' 后退一个索引，因为当前索引已经是下一个表了，由下个循环处理
                    ' 同时跳过表格
                    i = j - 1 + cellCount
                    
                    Exit For
                    
                 Else
                    
                    tableCaption = tableCaption + clearString(nextRange.Text)

                End If

            Next
            
            ' 至此，应该得到一个完整数据库表的声明，继续构造工具所需的格式
            If Len(tableDeclare) > 0 Then
                   
                aCSVTable = "##" + vbCrLf + TableName + "," + tableCaption + vbCrLf + tableDeclare + vbCrLf
                
                aCSVFile = aCSVFile + aCSVTable

            End If

        End If
    
    Next
    
    If Len(aCSVFile) > 0 Then
        ' 复制到粘贴板
        ' 需要配置 Microsoft Forms 2.0 Object Library
        ' fm20.dll
        Dim doClip As New DataObject
        doClip.SetText aCSVFile
        doClip.PutInClipboard
        
        MsgBox "已经复制到粘贴板"
        
    Else
    
        MsgBox "未找到任何数据库表"
    
    End If
    
    
End Sub

Sub 选中所有表格()
    Dim tempTable As Table
   
    Application.ScreenUpdating = False
   
    '判断文档是否被保护
    If ActiveDocument.ProtectionType = wdAllowOnlyFormFields Then
        MsgBox "文档已保护，此时不能选中多个表格！"
        Exit Sub
    End If
    '删除所有可编辑的区域
    ActiveDocument.DeleteAllEditableRanges wdEditorEveryone
    '添加可编辑区域
    For Each tempTable In ActiveDocument.Tables
        tempTable.Range.Editors.Add wdEditorEveryone
    Next
    '选中所有可编辑区域
    ActiveDocument.SelectAllEditableRanges wdEditorEveryone
    '删除所有可编辑的区域
    ActiveDocument.DeleteAllEditableRanges wdEditorEveryone
   
    Application.ScreenUpdating = True
   
End Sub

'=================多个文档，批量接受文档修订
Sub 批量接受多个docx的修订()
'
' acceptrevisions 宏
'
'
Dim myDialog As FileDialog, oDoc As Document
Dim oFile As Variant
On Error Resume Next
'定义一个文件夹选取对话框
Set myDialog = Application.FileDialog(msoFileDialogFilePicker)
With myDialog
    .Filters.Clear '清除所有文件筛选器中的项目
    .Filters.Add "所有WORD文件", "*.docx", 1 '增加筛选器的项目为所有Word文件
    .AllowMultiSelect = True '允许多项选择
    If .Show = -1 Then '确定
       For Each oFile In .SelectedItems '在所有选取项目中循环
       Set oDoc = Word.Documents.Open(fileName:=oFile, Visible:=False)
       oDoc.Revisions.AcceptAll
       oDoc.Close True
       Next
    End If
End With
End Sub



'=================多个文档，合并为一个文档（将多个文档放在同一路径下）
Sub 合并多个doc文档()
'
' combine 宏
'
'
Application.ScreenUpdating = False
MyPath = ActiveDocument.Path
MyName = Dir(MyPath & "\" & "*.doc")
i = 0
Do While MyName <> ""
If MyName <> ActiveDocument.Name Then
Set wb = Documents.Open(MyPath & "\" & MyName)
Selection.WholeStory
Selection.Copy
Windows(1).Activate
Selection.EndKey Unit:=wdLine
Selection.TypeParagraph
Selection.Paste
i = i + 1
wb.Close False
End If
MyName = Dir
Loop
Application.ScreenUpdating = True
End Sub

'=========================================一个文档中，多个图片大小修改
Sub 修改多个图片大小()
' setpicsize 宏
'
'
Dim n
Dim picwidth
Dim picheight
On Error Resume Next  '忽略错误
For n = 1 To ActiveDocument.InlineShapes.Count 'InlineShapes类型图片
picheight = ActiveDocument.InlineShapes(n).Height
picwidth = ActiveDocument.InlineShapes(n).Width
ActiveDocument.InlineShapes(n).Height = picheight * 0.8 '设置高度为0.6倍
ActiveDocument.InlineShapes(n).Width = picwidth * 0.8 '设置宽度为0.6倍
Next n
For n = 1 To ActiveDocument.Shapes.Count 'Shapes类型图片
picheight = ActiveDocument.Shapes(n).Height
picwidth = ActiveDocument.Shapes(n).Width
ActiveDocument.Shapes(n).Height = picheight * 0.8 '设置高度为0.6倍
ActiveDocument.Shapes(n).Width = picwidth * 0.8 '设置宽度为0.6倍
Next n
End Sub



'===========================所有的表格最后插入一列，并在最后一列加入列名
Sub 所有的表格最后插入一列()
'
' insertcolumn 宏，并在最后一列加入列名
'
'
If MsgBox("要为所有表格添加列吗？", vbYesNo + vbQuestion) = vbYes Then

For i = 1 To ActiveDocument.Tables.Count
    ActiveDocument.Tables(i).Columns.Add
    ActiveDocument.Tables(i).Columns.Last.Select
    Selection.TypeText Text:="实际结果"
Next

    MsgBox ("完成")

Else
    MsgBox ("任务取消")


End If

End Sub


'================调整所有表格的行高列宽

Sub 调整所有表格的宽度行高列宽()
'
' adjust_table_width 宏
'
'
For i = 1 To ActiveDocument.Tables.Count

  Dim thisTable As Word.Table
  Set thisTable = ActiveDocument.Tables(i)

  '调整表格宽度
  thisTable.Select
  'thisTable.PreferredWidth = CentimetersToPoints(16)
  
  
  thisTable.AutoFitBehavior (wdAutoFitContent) '根据内容自动调整表格
  thisTable.AutoFitBehavior (wdAutoFitWindow) '根据窗口自动调整表格
  'thisTable.range.ParagraphFormat.Alignment = wdAlignParagraphCenter '水平居中
  'thisTable.range.ParagraphFormat.Alignment = wdCellAlignVerticalCenter '垂直居中
  thisTable.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft '左对齐
  

    '调整行高
    'thisTable.Rows.Height = 40
    '调整行宽
    'thisTable.Columns.Width = 40
    
Next i
End Sub



Function isSimilar(table1 As Word.Table, table2 As Word.Table)

    isSimilar = False
    
    If table1.Columns.Count = table2.Columns.Count Then
        isSimilar = True
    End If
    
End Function

Sub 根据上一个表格调整列宽()


If ActiveDocument.Tables.Count < 1 Then
    Exit Sub
End If

'固定表宽度为100%
ActiveDocument.Tables(1).Select
ActiveDocument.Tables(1).PreferredWidthType = wdPreferredWidthPercent
ActiveDocument.Tables(1).PreferredWidth = 100

For i = 2 To ActiveDocument.Tables.Count Step 1


  Dim thisTable As Word.Table, previousTable As Word.Table
  
  Set previousTable = ActiveDocument.Tables(i - 1)
  Set thisTable = ActiveDocument.Tables(i)

  thisTable.Select

  '固定表宽度为100%
  thisTable.PreferredWidthType = wdPreferredWidthPercent
  thisTable.PreferredWidth = 100

  If isSimilar(previousTable, thisTable) Then
    
    
    
    'thisTable.range.ParagraphFormat.Alignment = wdAlignParagraphCenter '水平居中
    'thisTable.range.ParagraphFormat.Alignment = wdCellAlignVerticalCenter '垂直居中
    'thisTable.range.ParagraphFormat.Alignment = wdAlignParagraphLeft '左对齐
    
    
    For j = 1 To thisTable.Columns.Count Step 1
        
        thisTable.Columns(j).SetWidth previousTable.Columns(j).Width, wdAdjustNone
        
    Next j

    
  End If

    
Next i

End Sub


