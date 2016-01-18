Attribute VB_Name = "util"
'(C) 2015 Naoya Sawaguchi All Rights reserved.
Option Explicit

Function makeSheet()
Attribute makeSheet.VB_ProcData.VB_Invoke_Func = " \n14"
    Worksheets.Add After:=Sheets(Sheets.Count)
    
    makeSheet = Sheets.Count
End Function

Sub changeSheetName(ByVal index As Long, ByVal newName As String)
Attribute changeSheetName.VB_ProcData.VB_Invoke_Func = " \n14"
   Worksheets(index).name = newName
End Sub

Sub deleteSheet(ByVal index As Long)
  Application.DisplayAlerts = False
  Worksheets(index).Delete
  Application.DisplayAlerts = True
End Sub

Sub deleteOtherBookSheet(ByVal targetBookName As String, ByVal index As Long)
    Application.DisplayAlerts = False
    Workbooks(targetBookName).Sheets(index).Delete
    Application.DisplayAlerts = True
End Sub

Function copySheet(ByVal name As String)
    Sheets(name).Copy After:=Sheets(Sheets.Count)
    
    copySheet = Sheets.Count
End Function

Function copySheetToOtherBook(ByVal srcBookName As String, ByVal targetSheetName As String, ByVal dstBookName As String)
    Workbooks(srcBookName).Sheets(targetSheetName).Copy _
        After:=Workbooks(dstBookName).Sheets(Workbooks(dstBookName).Sheets.Count)
    
    copySheetToOtherBook = Workbooks(dstBookName).Sheets.Count
End Function

Function isStrEmpty(ByVal v As String)
    isStrEmpty = (StrComp(v, "") = 0)
End Function

Sub setText(ByVal sheetIdx As Long, ByVal shapeName As String, ByVal txt As String)
    On Error Resume Next
    Sheets(sheetIdx).Shapes(shapeName).TextFrame.Characters.text = txt
    
End Sub

Sub deleteShape(ByVal sheetIdx As Long, ByVal shapeName As String)
    Sheets(sheetIdx).Shapes(shapeName).Delete
End Sub
 
 'return new workbook's name
Function openNewBook()
    Workbooks.Add
    openNewBook = ActiveWorkbook.name
End Function
