---
title: Worksheet Object Events
keywords: vbaxl10.chm5206017
f1_keywords:
- vbaxl10.chm5206017
ms.prod: excel
ms.assetid: 512e329c-92f6-a8e0-8564-b3ba57e8c296
ms.date: 06/08/2017
---


# Worksheet Object Events

Events on sheets are enabled by default. To view the event procedures for a sheet, right-click the sheet tab and click  **View Code** on the shortcut menu. Select one of the following events from the **Procedure** drop-down list box.



[Activate](worksheet-activate-event-excel.md) | 
[BeforeDoubleClick](worksheet-beforedoubleclick-event-excel.md) | 
[BeforeRightClick](worksheet-beforerightclick-event-excel.md) | 
[Calculate](worksheet-calculate-event-excel.md) | 
[Change](worksheet-change-event-excel.md) | 
[Deactivate](worksheet-deactivate-event-excel.md) | 
[FollowHyperlink](worksheet-followhyperlink-event-excel.md) | 
[PivotTableUpdate](worksheet-pivottableupdate-event-excel.md) | 
[SelectionChange](worksheet-selectionchange-event-excel.md)

Worksheet-level events occur when a worksheet is activated, when the user changes a worksheet cell, or when the PivotTable changes. The following example adjusts the size of columns A through F whenever the worksheet is recalculated.




```vb
Private Sub Worksheet_Calculate() 
    Columns("A:F").AutoFit 
End Sub
```

Some events can be used to substitute an action for the default application behavior, or to make a small change to the default behavior. The following example traps the right-click event and adds a new menu item to the shortcut menu for cells B1:B10.



```vb
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, _ 
        Cancel As Boolean) 
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


