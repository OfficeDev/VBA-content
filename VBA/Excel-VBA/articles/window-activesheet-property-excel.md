---
title: Window.ActiveSheet Property (Excel)
keywords: vbaxl10.chm356079
f1_keywords:
- vbaxl10.chm356079
ms.prod: excel
api_name:
- Excel.Window.ActiveSheet
ms.assetid: 44e4fd8d-45bd-5626-66db-107fb451b73f
ms.date: 06/08/2017
---


# Window.ActiveSheet Property (Excel)

Returns an object that represents the active sheet (the sheet on top) in the active workbook or in the specified window or workbook. Returns  **Nothing** if no sheet is active.


## Syntax

 _expression_ . **ActiveSheet**

 _expression_ A variable that represents a **Window** object.


## Remarks

If you don't specify an object qualifier, this property returns the active sheet in the active workbook.

If a workbook appears in more than one window, the  **ActiveSheet** property may be different in different windows.


## Example

This example displays the name of the active sheet.


```vb
MsgBox "The name of the active sheet is " &; ActiveSheet.Name
```


## See also


#### Concepts


[Window Object](window-object-excel.md)

