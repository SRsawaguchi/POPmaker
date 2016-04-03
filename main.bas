Attribute VB_Name = "main"
'(C) 2015 Naoya Sawaguchi All Rights reserved.
Option Explicit

'�V�[�g��̍��ږ�
Const HEADING_MESSAGE = "��P����"
Const SUB_MESSAGE = "��Q����"
Const MAKER_NAME = "���[�J�[��"
Const ITEM_NAME_1 = "���i��"
Const ITEM_NAME_2 = "���Ȃ�"
Const UNIT = "����"
Const PRICE = "�P��"
Const PRINT_COUNT = "�������"

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

    'ProgressDialog�̕\��
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
    
    '�e���v���[�g���폜
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
    '�����I���̃A�i�E���X(���܂ɍs���Ȃ�����O�̂��߂Q�����Ƃ�
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
    
    '�u���Ȃǁv�̍��ڂ͈����p����Ȃ�
    pop.setItemName2 (target.getItemName2())
    
    If Not isStrEmpty(target.getUnit()) Then
        pop.setUnit (target.getUnit())
    End If
    
    If Not isStrEmpty(target.getPrice()) Then
        pop.setPrice (target.getPrice())
    End If
    
    '������������͂Ȃ�-1���ݒ肳��Ă���B
    '�����͂̏ꍇ�A�����p��
    If target.getPrintCount > -1 Then
        pop.setPrintCount (target.getPrintCount())
    End If
    
    Set mergeDefaultPop = pop
End Function

Sub onSettingButtonClick()
    Sheet2.Activate
End Sub
