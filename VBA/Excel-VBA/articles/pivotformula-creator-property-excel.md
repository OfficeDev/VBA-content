---
title: PivotFormula.Creator Property (Excel)
keywords: vbaxl10.chm230074
f1_keywords:
- vbaxl10.chm230074
ms.prod: excel
api_name:
- Excel.PivotFormula.Creator
ms.assetid: d3d302ec-3f9a-7969-bfbe-51e56680cce5
ms.date: 06/08/2017
---


# PivotFormula.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **PivotFormula** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[PivotFormula Object](pivotformula-object-excel.md)

