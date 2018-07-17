---
title: PivotLine.Creator Property (Excel)
keywords: vbaxl10.chm763074
f1_keywords:
- vbaxl10.chm763074
ms.prod: excel
api_name:
- Excel.PivotLine.Creator
ms.assetid: 9f68797c-1817-eff5-3b5e-17371961fc2c
ms.date: 06/08/2017
---


# PivotLine.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **PivotLine** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[PivotLine Object](pivotline-object-excel.md)

