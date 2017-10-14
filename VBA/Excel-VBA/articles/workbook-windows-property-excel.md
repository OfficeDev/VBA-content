---
title: Workbook.Windows Property (Excel)
keywords: vbaxl10.chm199165
f1_keywords:
- vbaxl10.chm199165
ms.prod: excel
api_name:
- Excel.Workbook.Windows
ms.assetid: 2352d6c9-720e-b58d-6e7c-049bf21a090d
ms.date: 06/08/2017
---


# Workbook.Windows Property (Excel)

Returns a  **[Windows](windows-object-excel.md)** collection that represents all the windows in the specified workbook. Read-only **Windows** object.


## Syntax

 _expression_ . **Windows**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

Using this property without an object qualifier is equivalent to using  `Application.Windows`.

This property returns a collection of both visible and hidden windows.


## Example

This example names window one in the active workbook "Consolidated Balance Sheet." This name is then used as the index to the  **Windows** collection.


```vb
ActiveWorkbook.Windows(1).Caption = "Consolidated Balance Sheet" 
ActiveWorkbook.Windows("Consolidated Balance Sheet") _ 
 .ActiveSheet.Calculate
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

