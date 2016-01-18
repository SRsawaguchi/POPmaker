VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PopSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'(C) 2015 Naoya Sawaguchi All Rights reserved.
Option Explicit

'図形参照用のプレフィックス
Const HEADING_MESSAGE_PREFIX = "HeadingMessage"
Const SUB_MESSAGE_PREFIX = "SubMessage"
Const MAKER_NAME_PREFIX = "MakerName"
Const UNIT_PREFIX = "Unit"
Const ITEM_NAME_PREFIX = "ItemName"
Const PRICE_PREFIX = "Price"
Const GROUP_TAX_YEN = "Group" '税抜き、円の画像グループ

Private p_maxPopCount As Long
Private p_templateSheetName As String
Private p_sheetName As String
Private p_index As Long
Private p_pops() As PopData
Private p_fillPointer As Long
Private p_printed As Boolean

Public Sub init(ByVal sheetName As String, ByVal templateSheetName As String, ByVal maxPopCount As Long)
    p_sheetName = sheetName
    p_maxPopCount = maxPopCount
    p_fillPointer = 0
    p_printed = False
    p_templateSheetName = templateSheetName
    ReDim p_pops(p_maxPopCount - 1)
End Sub

Public Function addPopData(ByVal pop As PopData)
    Set p_pops(p_fillPointer) = pop
    p_fillPointer = p_fillPointer + 1
End Function

Public Sub printPop()
    Dim pop As PopData
    Dim i As Long
    
    p_index = copySheet(p_templateSheetName)
    Call changeSheetName(p_index, p_sheetName)
    
    For i = 0 To UBound(p_pops)
            
        If i > p_fillPointer - 1 Then
            '入力されていない分は空文字列を挿入し、「税込み、円」を削除してトナーを節約する。
            Call setText(p_index, HEADING_MESSAGE_PREFIX & i, "")
            Call setText(p_index, SUB_MESSAGE_PREFIX & i, "")
            Call setText(p_index, MAKER_NAME_PREFIX & i, "")
            Call setText(p_index, ITEM_NAME_PREFIX & i, "")
            Call setText(p_index, UNIT_PREFIX & i, "")
            Call setText(p_index, PRICE_PREFIX & i, "")
            Call deleteShape(p_index, GROUP_TAX_YEN & i)
        Else
            Set pop = p_pops(i)
            Call setText(p_index, HEADING_MESSAGE_PREFIX & i, pop.getHeadingMessage())
            Call setText(p_index, SUB_MESSAGE_PREFIX & i, pop.getSubMessage())
            Call setText(p_index, MAKER_NAME_PREFIX & i, pop.getMakerName())
            Call setText(p_index, ITEM_NAME_PREFIX & i, pop.getItemName())
            Call setText(p_index, UNIT_PREFIX & i, pop.getUnit())
            Call setText(p_index, PRICE_PREFIX & i, pop.getPrice())
        End If
    Next
    
    p_printed = True
End Sub

Public Function getMaxPopCount()
    getMaxPopCount = p_maxPopCount
End Function

Public Function getSheetName()
    getSheetName = p_sheetName
End Function

Public Function getIndex()
    getIndex = p_index
End Function

Public Function getCurrentPopCount()
    getCurrentPopCount = p_fillPointer
End Function

Public Function hasPrinted()
    hasPrinted = p_printed
End Function
