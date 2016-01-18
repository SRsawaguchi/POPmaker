VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'(C) 2015 Naoya Sawaguchi All Rights reserved.
Option Explicit

Public Function getPopDataTable()
    Set getPopDataTable = Range(Sheet1.Cells(4, 2), Sheet1.Cells(54, 8))
End Function

Public Function getPopDataTableWithoutTitle()
    Set getPopDataTableWithoutTitle = Range(Sheet1.Cells(5, 2), Sheet1.Cells(54, 8))
End Function

Sub onClickTableSelectButton()
    Dim r As Range
    Set r = getPopDataTableWithoutTitle
    r.Select
End Sub

Sub onClickHowToPrintButton()
    Sheet5.Activate
End Sub

Sub onClickHowToButton()
    Sheet8.Activate
End Sub

Sub onClickAboutButton()
    AboutDialog.Show
    
End Sub
