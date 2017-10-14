---
title: PivotLines.Creator Property (Excel)
keywords: vbaxl10.chm765074
f1_keywords:
- vbaxl10.chm765074
ms.prod: excel
api_name:
- Excel.PivotLines.Creator
ms.assetid: 090d80a7-f0e8-4b5c-4095-84b9304f4c3f
ms.date: 06/08/2017
---


# PivotLines.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **PivotLines** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[PivotLines Object](pivotlines-object-excel.md)

