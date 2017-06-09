---
title: IconSet.Creator Property (Excel)
keywords: vbaxl10.chm817074
f1_keywords:
- vbaxl10.chm817074
ms.prod: excel
api_name:
- Excel.IconSet.Creator
ms.assetid: 32801791-c2d6-04d2-e93d-b6583728ced8
ms.date: 06/08/2017
---


# IconSet.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents an **IconSet** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[IconSet Object](iconset-object-excel.md)

