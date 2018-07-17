---
title: UniqueValues.Creator Property (Excel)
keywords: vbaxl10.chm825074
f1_keywords:
- vbaxl10.chm825074
ms.prod: excel
api_name:
- Excel.UniqueValues.Creator
ms.assetid: d710b769-8c9b-12f9-ff31-77d4bb14bf64
ms.date: 06/08/2017
---


# UniqueValues.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **UniqueValues** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[UniqueValues Object](uniquevalues-object-excel.md)

