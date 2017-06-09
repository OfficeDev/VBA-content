---
title: Windows.BreakSideBySide Method (Excel)
keywords: vbaxl10.chm354079
f1_keywords:
- vbaxl10.chm354079
ms.prod: excel
api_name:
- Excel.Windows.BreakSideBySide
ms.assetid: be32b6a4-5541-8c4b-ef24-cf34c9035f1c
ms.date: 06/08/2017
---


# Windows.BreakSideBySide Method (Excel)

Ends side-by-side mode if two windows are in side-by-side mode. Returns a  **Boolean** value that represents whether the method was successful.


## Syntax

 _expression_ . **BreakSideBySide**

 _expression_ A variable that represents a **Windows** object.


### Return Value

Boolean


## Example

The following example ends side-by-side mode.


```vb
Sub CloseSideBySide() 
 
 ActiveWorkbook.Windows.BreakSideBySide 
 
End Sub
```


## See also


#### Concepts


[Windows Object](windows-object-excel.md)

