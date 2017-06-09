---
title: Names.Creator Property (Excel)
keywords: vbaxl10.chm487074
f1_keywords:
- vbaxl10.chm487074
ms.prod: excel
api_name:
- Excel.Names.Creator
ms.assetid: 7584df14-1683-a80d-ec09-2354bdb4e71d
ms.date: 06/08/2017
---


# Names.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Names** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Names Object](names-object-excel.md)

