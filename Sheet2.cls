VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'(C) 2015 Naoya Sawaguchi All Rights reserved.
Option Explicit

Public Enum PaperSize
    DIV_2 = 1 '横２枚切り（合同企画POPなど)
    DIV_4_TBL = 2 'パンのPOP
End Enum

Sub onHomeButtonClick() '「POP作成画面へ」ボタンクリック時
    Sheet1.Activate
End Sub

Sub onVisibleButtonClick()
    Dim templates() As String
    Dim name As Variant
    templates = getTemplateArray()
    
    For Each name In templates
        Sheets(name).Visible = True
    Next
End Sub

Sub onHiddenButtonClick()
    Dim templates() As String
    Dim name As Variant
    templates = getTemplateArray()
    
    For Each name In templates
        Sheets(name).Visible = False
    Next
End Sub

Function getPopDefaultTable()
    Dim r As Range
    Set r = Sheet2.Range(Cells(8, 2), Cells(9, 8))
    Set getPopDefaultTable = r
End Function
