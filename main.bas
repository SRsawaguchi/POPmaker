Attribute VB_Name = "main"
'(C) 2015 Naoya Sawaguchi All Rights reserved.
Option Explicit

Dim g_outBookName As String

Sub onMakeButtonClick()
    Dim i, maxPop As Long
    Dim orModeEnable As Boolean
    Dim popPaperName As String
    Dim macroBookName As String
    Dim targetBookName As String
    Dim table As Range
    Dim defaultTable As Range
    Dim pop As PopData
    Dim popFactory As PopDataFactory
    Dim sheetFactory As PopSheetFactory
    
    macroBookName = ThisWorkbook.name

    Set defaultTable = Sheet2.getPopDefaultTable()

    Set table = Sheet1.getPopDataTable()
    Set popFactory = New PopDataFactory
    Call popFactory.init(table, defaultTable)
    
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
    
    For i = 2 To table.Rows.Count - 1
        Set pop = popFactory.make(i)
        
        If orModeEnable Then
            Call popFactory.setDefaultPop(pop)
        End If
        
        If Not isStrEmpty(pop.getItemName()) Then
            Call sheetFactory.pushPopData(pop)
        End If
        
        Call ProgressDialog.onProgressUpdate(i)
    Next
    
    Call sheetFactory.closeFactory
    
    For i = 1 To Workbooks(targetBookName).Sheets.Count
        If Not Workbooks(targetBookName).Sheets(1).name Like "*pop*" Then
            If Workbooks(targetBookName).Sheets.Count > 1 Then
                deleteSheet (1)
            End If
        End If
    Next
    
    
    '処理終了のアナウンス
    Application.ScreenUpdating = True
    Workbooks(macroBookName).Activate
    Sheets(popPaperName).Visible = False
    Call ProgressDialog.onProcessFinished(targetBookName)
End Sub

Sub onSettingButtonClick()
    Sheet2.Activate
End Sub
