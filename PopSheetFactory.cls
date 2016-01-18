VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PopSheetFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'(C) 2015 Naoya Sawaguchi All Rights reserved.
Option Explicit

Const SHEET_NAME_PREFIX = "pop"

Private p_curSheet As PopSheet
Private p_templateSheetName As String
Private p_sheetCount As Long
Private p_maxPopCount As Long
Private p_closed As Boolean


Public Sub init(ByVal templateSheetName As String, ByVal maxPopCount As Long)
    p_templateSheetName = templateSheetName
    p_maxPopCount = maxPopCount
    p_sheetCount = 0
    p_closed = False
    Set p_curSheet = makeNewPopSheet()
End Sub

Private Function makeNewPopSheet()
    Dim newPopSheet As PopSheet
    Set newPopSheet = New PopSheet
    Call newPopSheet.init(SHEET_NAME_PREFIX & p_sheetCount, p_templateSheetName, p_maxPopCount)
    
    Set makeNewPopSheet = newPopSheet
End Function

Public Sub pushPopData(ByVal pop As PopData)
    Dim cnt As Long
    
    If p_curSheet.getCurrentPopCount() >= p_maxPopCount Then
        Call p_curSheet.printPop
        p_sheetCount = p_sheetCount + 1
        Set p_curSheet = makeNewPopSheet()
    End If
    
    If pop.getPrintCount > 0 Then
        Call p_curSheet.addPopData(pop)
        cnt = pop.getPrintCount()
        cnt = cnt - 1
        pop.setPrintCount (cnt)
        
        Call Me.pushPopData(pop) 'Ä‹A
    End If
End Sub

Public Sub closeFactory()
    If p_curSheet.getCurrentPopCount > 0 And p_curSheet.hasPrinted() = False Then
        Call p_curSheet.printPop
    End If
    
    p_closed = True
End Sub


Public Function hasClosed()
    hasClosed = p_closed
End Function
