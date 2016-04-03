Attribute VB_Name = "main"
'(C) 2015 Naoya Sawaguchi All Rights reserved.
Option Explicit

'シート上の項目名
Const HEADING_MESSAGE = "第１文言"
Const SUB_MESSAGE = "第２文言"
Const MAKER_NAME = "メーカー名"
Const ITEM_NAME_1 = "商品名"
Const ITEM_NAME_2 = "味など"
Const UNIT = "数量"
Const PRICE = "単価"
Const PRINT_COUNT = "印刷枚数"

Dim g_outBookName As String

Sub onMakeButtonClick()
    Dim i, maxPop As Long
    Dim orModeEnable As Boolean
    Dim popPaperName As String
    Dim macroBookName As String
    Dim targetBookName As String
    Dim table As Range
    Dim defaultTable As Range
    Dim defaultPop  As PopData
    Dim newPop As PopData
    Dim readPop As PopData
    Dim sheetFactory As PopSheetFactory
    
    macroBookName = ThisWorkbook.name

    Set defaultTable = Sheet2.getPopDefaultTable()
    Set table = Sheet1.getPopDataTable()
    
    popPaperName = getSelectedTemplateName()
    maxPop = getMaxPopCount()
    orModeEnable = overrideModeEnable()
    
    Sheets(popPaperName).Visible = True
    Set sheetFactory = New PopSheetFactory
    Application.ScreenUpdating = False
    targetBookName = openNewBook()

    'ProgressDialogの表示
    Call ProgressDialog.onProcessStart(macroBookName, 1, table.Rows.Count)
    DoEvents
    i = copySheetToOtherBook(macroBookName, popPaperName, targetBookName)
    Workbooks(targetBookName).Activate
    Call sheetFactory.init(popPaperName, maxPop)
    
    Set defaultPop = makePopData(defaultTable, 2)
    For i = 2 To table.Rows.Count - 1
        Set readPop = makePopData(table, i)
        
        If Not (isStrEmpty(readPop.getItemName1()) And isStrEmpty(readPop.getItemName2())) Then
            Set newPop = mergeDefaultPop(defaultPop, readPop)
            
            If orModeEnable Then
                Set defaultPop = newPop.clone
            End If
                
            Call sheetFactory.pushPopData(newPop)
        End If
        
        Call ProgressDialog.onProgressUpdate(i)
    Next
    
    Call sheetFactory.closeFactory
    
    'テンプレートを削除
    For i = 1 To Workbooks(targetBookName).Sheets.Count
        If Not Workbooks(targetBookName).Sheets(1).name Like "*pop*" Then
            If Workbooks(targetBookName).Sheets.Count > 1 Then
                deleteSheet (1)
            End If
        End If
    Next
    
    Application.ScreenUpdating = True
    Workbooks(macroBookName).Activate
    Sheets(popPaperName).Visible = False
    Call ProgressDialog.onProcessFinished(targetBookName)
    '処理終了のアナウンス(たまに行かないから念のため２回やっとく
    Application.ScreenUpdating = True
End Sub

Private Function findData(ByVal table As Range, ByVal title As String, ByVal idx As Long)
    findData = WorksheetFunction.HLookup(title, table, idx, False)
End Function

Private Function makePopData(ByVal table As Range, ByVal idx As Long)
    Dim pop As New PopData
    Dim s_cnt As String
    Dim item_name As String
    Dim l_cnt As Long
             
    pop.setHeadingMessage (findData(table, HEADING_MESSAGE, idx))
    pop.setSubMessage (findData(table, SUB_MESSAGE, idx))
    pop.setMakerName (findData(table, MAKER_NAME, idx))
    pop.setItemName1 (findData(table, ITEM_NAME_1, idx))
    pop.setItemName2 (findData(table, ITEM_NAME_2, idx))
    pop.setUnit (findData(table, UNIT, idx))
    pop.setPrice (findData(table, PRICE, idx))
    
    s_cnt = findData(table, PRINT_COUNT, idx)
    If Not isStrEmpty(s_cnt) Then
        l_cnt = Val(s_cnt)
        pop.setPrintCount (findData(table, PRINT_COUNT, idx))
    Else
        pop.setPrintCount (-1)
    End If
    
    Set makePopData = pop
End Function

Public Function mergeDefaultPop(default As PopData, target As PopData)
    Dim pop As PopData
    
    Set pop = default.clone
    
    If Not isStrEmpty(target.getHeadingMessage()) Then
        pop.setHeadingMessage (target.getHeadingMessage())
    End If
    
    If Not isStrEmpty(target.getSubMessage()) Then
        pop.setSubMessage (target.getSubMessage())
    End If
    
    If Not isStrEmpty(target.getMakerName()) Then
        pop.setMakerName (target.getMakerName())
    End If
    
    If Not isStrEmpty(target.getItemName1()) Then
        pop.setItemName1 (target.getItemName1())
    End If
    
    '「味など」の項目は引き継がれない
    pop.setItemName2 (target.getItemName2())
    
    If Not isStrEmpty(target.getUnit()) Then
        pop.setUnit (target.getUnit())
    End If
    
    If Not isStrEmpty(target.getPrice()) Then
        pop.setPrice (target.getPrice())
    End If
    
    '印刷枚数未入力なら-1が設定されている。
    '未入力の場合、引き継ぎ
    If target.getPrintCount > -1 Then
        pop.setPrintCount (target.getPrintCount())
    End If
    
    Set mergeDefaultPop = pop
End Function

Sub onSettingButtonClick()
    Sheet2.Activate
End Sub
