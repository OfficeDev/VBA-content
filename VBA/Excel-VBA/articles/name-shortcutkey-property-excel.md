---
title: Name.ShortcutKey Property (Excel)
keywords: vbaxl10.chm490081
f1_keywords:
- vbaxl10.chm490081
ms.prod: excel
api_name:
- Excel.Name.ShortcutKey
ms.assetid: ff763568-4c18-9414-45a7-bcf75b597261
ms.date: 06/08/2017
---


# Name.ShortcutKey Property (Excel)

Returns or sets the shortcut key for a name defined as a custom Microsoft Excel 4.0 macro command. Read/write  **String** .


## Syntax

 _expression_ . **ShortcutKey**

 _expression_ A variable that represents a **Name** object.


## Example

This example sets the shortcut key for name one in the active workbook. The example should be run on a workbook in which name one refers to a Microsoft Excel 4.0 command macro.


```vb
ActiveWorkbook.Names(1).ShortcutKey = "K"
```


## See also


#### Concepts


[Name Object](name-object-excel.md)

