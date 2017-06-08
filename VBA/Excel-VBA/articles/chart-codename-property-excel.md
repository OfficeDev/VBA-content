---
title: Chart.CodeName Property (Excel)
keywords: vbaxl10.chm148076
f1_keywords:
- vbaxl10.chm148076
ms.prod: excel
api_name:
- Excel.Chart.CodeName
ms.assetid: 88872dad-53b2-580a-9bbc-6a29066352a6
ms.date: 06/08/2017
---


# Chart.CodeName Property (Excel)

Returns the code name for the object. Read-only  **String** .


## Syntax

 _expression_ . **CodeName**

 _expression_ An expression that returns a **Chart** object.


## Remarks

The value that you see in the cell to the right of  **(Name)** in the **Properties** window is the code name of the selected object. At design time, you can change the code name of an object by changing this value. You cannot programmatically change this property at run time.

The code name for an object can be used in place of an expression that returns the object. For example, if the code name for worksheet one is "Sheet1", the following expressions are identical:




```vb
Worksheets(1).Range("a1") 
Sheet1.Range("a1")
```

It's possible for the sheet name to be different from the code name. When you create a sheet, the sheet name and code name are the same, but changing the sheet name doesn't change the code name, and changing the code name (using the  **Properties** window in the Visual Basic Editor) doesn't change the sheet name.


## Example

This example displays the code name for worksheet one.


```vb
MsgBox Worksheets(1).CodeName
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

