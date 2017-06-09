---
title: Workbook.AddinInstall Event (Excel)
keywords: vbaxl10.chm503080
f1_keywords:
- vbaxl10.chm503080
ms.prod: excel
api_name:
- Excel.Workbook.AddinInstall
ms.assetid: 671117b2-590e-9d6f-29ae-5f0bf30d4e99
ms.date: 06/08/2017
---


# Workbook.AddinInstall Event (Excel)

Occurs when the workbook is installed as an add-in


## Syntax

 _expression_ . **AddinInstall**

 _expression_ A variable that represents a **Workbook** object.


### Return Value

Nothing


## Example

This example adds a control to the standard toolbar when the workbook is installed as an add-in.


```vb
Private Sub Workbook_AddinInstall() 
 With Application.Commandbars("Standard").Controls.Add 
 .Caption = "The AddIn's menu item" 
 .OnAction = "'ThisAddin.xls'!Amacro" 
 End With End Sub 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

