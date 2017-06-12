---
title: DisplayFormat.FormulaHidden Property (Excel)
keywords: vbaxl10.chm893078
f1_keywords:
- vbaxl10.chm893078
ms.prod: excel
api_name:
- Excel.DisplayFormat.FormulaHidden
ms.assetid: 3db0fd6b-da1b-f19a-e859-a949b5f4d2b3
ms.date: 06/08/2017
---


# DisplayFormat.FormulaHidden Property (Excel)

Returns a value that indicates if the formula of the associated  **[Range](range-object-excel.md)** object is hidden when the worksheet is protected as it is displayed in the current user interface. Read-only.


## Syntax

 _expression_ . **FormulaHidden**

 _expression_ A variable that represents a **[DisplayFormat](displayformat-object-excel.md)** object.


### Return Value

Variant


## Remarks

Returns  **True** if the formula is hidden when the worksheet is protected, **Null** if the range contains some cells with **FormulaHidden** equal to **True** and some cells with **FormulaHidden** equal to **False** .


## See also


#### Concepts


[DisplayFormat Object](displayformat-object-excel.md)

