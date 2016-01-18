VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PopData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'(C) 2015 Naoya Sawaguchi All Rights reserved.
Option Explicit

Private p_headingMessage As String 'お買い得品
Private p_subMessage As String '1日～3日
Private p_makerName As String '日清食品
Private p_itemName As String 'カップヌードル
Private p_unit As String '１コ
Private p_price As String '113円
Private p_printCount As Long '印刷枚数


Public Function getHeadingMessage()
    getHeadingMessage = p_headingMessage
End Function

Public Function setHeadingMessage(ByVal msg As String)
    If isStrEmpty(msg) = False Then
        p_headingMessage = msg
    End If
End Function

Public Function getSubMessage()
    getSubMessage = p_subMessage
End Function

Public Function setSubMessage(ByVal msg As String)
    If isStrEmpty(msg) = False Then
        p_subMessage = msg
    End If
End Function

Public Function getMakerName()
    getMakerName = p_makerName
End Function

Public Function setMakerName(ByVal name As String)
    If isStrEmpty(name) = False Then
        p_makerName = name
    End If
End Function

Public Function getItemName()
    getItemName = p_itemName
End Function

Public Function setItemName(ByVal name As String)
        p_itemName = name
End Function

Public Function getUnit()
    getUnit = p_unit
End Function

Public Function setUnit(ByVal u As String)
    If isStrEmpty(u) = False Then
        p_unit = u
    End If
End Function

Public Function getPrice()
    getPrice = p_price
End Function

Public Function setPrice(ByVal p As String)
    If isStrEmpty(p) = False Then
        p_price = p
    End If
End Function

Public Function getPrintCount()
    getPrintCount = p_printCount
End Function

Public Function setPrintCount(ByVal pc As Long)
    p_printCount = pc
End Function

Public Function clone()
    Dim pop As PopData
    Set pop = New PopData
    
    pop.setHeadingMessage (p_headingMessage)
    pop.setSubMessage (p_subMessage)
    pop.setMakerName (p_makerName)
    pop.setItemName (p_itemName)
    pop.setUnit (p_unit)
    pop.setPrice (p_price)
    pop.setPrintCount (p_printCount)
    
    Set clone = pop
End Function
