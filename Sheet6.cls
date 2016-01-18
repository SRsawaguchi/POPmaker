VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'(C) 2015 Naoya Sawaguchi All Rights reserved.

Public Function getReleaseVersion()
    getReleaseVersion = Sheet6.Cells(1, 4).Value
End Function
