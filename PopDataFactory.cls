VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PopDataFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'(C) 2015 Naoya Sawaguchi All Rights reserved.
Option Explicit

'シート上の項目名
Const HEADING_MESSAGE = "第１文言"
Const SUB_MESSAGE = "第２文言"
Const MAKER_NAME = "メーカー名"
Const ITEM_NAME = "商品名"
Const UNIT = "数量"
Const PRICE = "単価"
Const PRINT_COUNT = "印刷枚数"

Private p_table As Range
Private p_defaultTable As Range
Private p_defaultPop As PopData

Public Sub init(ByVal table As Range, ByVal defaultTable As Range)
    Set p_table = table
    Set p_defaultTable = defaultTable
    Set p_defaultPop = makePopData(defaultTable, 2)
End Sub

Private Function makePopData(ByVal table As Range, ByVal idx As Long)
    Dim pop As PopData
    Dim s_cnt As String
    Dim l_cnt As Long
    
    If p_defaultPop Is Nothing Then
        Set pop = New PopData
    Else
        Set pop = p_defaultPop.clone()
    End If
           
    pop.setHeadingMessage (findData(table, HEADING_MESSAGE, idx))
    pop.setSubMessage (findData(table, SUB_MESSAGE, idx))
    pop.setMakerName (findData(table, MAKER_NAME, idx))
    pop.setItemName (findData(table, ITEM_NAME, idx))
    pop.setUnit (findData(table, UNIT, idx))
    pop.setPrice (findData(table, PRICE, idx))
    
    s_cnt = findData(table, PRINT_COUNT, idx)
    If Not isStrEmpty(s_cnt) Then
        l_cnt = Val(s_cnt)
        pop.setPrintCount (findData(table, PRINT_COUNT, idx))
    End If
    
    Set makePopData = pop
End Function

Private Function findData(ByVal table As Range, ByVal title As String, ByVal idx As Long)
    findData = WorksheetFunction.HLookup(title, table, idx, False)
End Function


Public Function make(ByVal idx As Long)
    Set make = makePopData(p_table, idx)
End Function

Public Function getTable()
    getTable = p_table
End Function

Public Function getDefaultPop()
    Set getDefaultPop = p_defaultPop
End Function

Public Function setDefaultPop(ByVal pop As PopData)
    Set p_defaultPop = pop.clone
    Call p_defaultPop.setItemName("")
End Function
