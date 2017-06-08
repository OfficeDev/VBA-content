---
title: Worksheet.BeforeRightClick Event (Excel)
keywords: vbaxl10.chm502075
f1_keywords:
- vbaxl10.chm502075
ms.prod: excel
api_name:
- Excel.Worksheet.BeforeRightClick
ms.assetid: 0263dd09-1648-d3c4-007e-15ef7b82092a
ms.date: 06/08/2017
---


# Worksheet.BeforeRightClick Event (Excel)

Occurs when a worksheet is right-clicked, before the default right-click action.


## Syntax

 _expression_ . **BeforeRightClick**( **_Target_** , **_Cancel_** )

 _expression_ A variable that represents a **Worksheet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Target_|Required| **Range**|The cell nearest to the mouse pointer when the right-click occurs.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the default right-click action doesn't occur when the procedure is finished.|

## Remarks

Like other worksheet events, this event doesn't occur if you right-click while the pointer is on a shape or a command bar (a toolbar or menu bar).


## Example

This example adds a new menu item to the shortcut menu for cells B1:B10.


```vb
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, _ 
 Cancel As Boolean) 
 Dim icbc As Object 
 For Each icbc In Application.CommandBars("cell").Controls 
 If icbc.Tag = "brccm" Then icbc.Delete 
 Next icbc 
 If Not Application.Intersect(Target, Range("b1:b10")) _ 
 Is Nothing Then 
 With Application.CommandBars("cell").Controls _ 
 .Add(Type:=msoControlButton, before:=6, _ 
 temporary:=True) 
 .Caption = "New Context Menu Item" 
 .OnAction = "MyMacro" 
 .Tag = "brccm" 
 End With 
 End If 
End Sub
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

