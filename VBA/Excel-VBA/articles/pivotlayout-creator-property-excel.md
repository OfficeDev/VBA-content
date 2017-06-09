---
title: PivotLayout.Creator Property (Excel)
keywords: vbaxl10.chm663074
f1_keywords:
- vbaxl10.chm663074
ms.prod: excel
api_name:
- Excel.PivotLayout.Creator
ms.assetid: 0cbe7f15-997c-c395-879d-64aa43dff05d
ms.date: 06/08/2017
---


# PivotLayout.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **PivotLayout** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[PivotLayout Object](pivotlayout-object-excel.md)

