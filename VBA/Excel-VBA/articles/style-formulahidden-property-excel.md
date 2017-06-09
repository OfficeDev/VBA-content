---
title: Style.FormulaHidden Property (Excel)
keywords: vbaxl10.chm177078
f1_keywords:
- vbaxl10.chm177078
ms.prod: excel
api_name:
- Excel.Style.FormulaHidden
ms.assetid: 7b36f86b-2f88-3fb4-173e-cca7e747a195
ms.date: 06/08/2017
---


# Style.FormulaHidden Property (Excel)

Returns or sets a  **Boolean** value that indicates if the formula will be hidden when the worksheet is protected.


## Syntax

 _expression_ . **FormulaHidden**

 _expression_ A variable that represents a **Style** object.


## Remarks

Don't confuse this property with the  **[Hidden](range-hidden-property-excel.md)** property. The formula will not be hidden if the workbook is protected and the worksheet is not, but only if the worksheet is protected.


## See also


#### Concepts


[Style Object](style-object-excel.md)

