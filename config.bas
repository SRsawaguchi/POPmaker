Attribute VB_Name = "Config"
'(C) 2015 Naoya Sawaguchi All Rights reserved.
Option Explicit

Const CONF_SHEET_NAME = "íËêîÅEê›íËä«óù"

Function overrideModeEnable()
    overrideModeEnable = Sheets(CONF_SHEET_NAME).Cells(6, 12).Value
End Function

Function getSelectedTemplateName()
    getSelectedTemplateName = Sheets(CONF_SHEET_NAME).Cells(6, 5).Value
End Function

Function getMaxPopCount()
    getMaxPopCount = Sheets(CONF_SHEET_NAME).Cells(6, 6).Value
End Function

Function getTemplateArray()
    Dim max, i As Long
    Dim arr() As String
    Dim name As String
    
    max = Sheets(CONF_SHEET_NAME).Cells(6, 2).Value - 1
    ReDim arr(max) As String
    
    For i = 0 To max
        arr(i) = Sheets(CONF_SHEET_NAME).Cells(8 + i, 2).Value
    Next
    
    getTemplateArray = arr
End Function
