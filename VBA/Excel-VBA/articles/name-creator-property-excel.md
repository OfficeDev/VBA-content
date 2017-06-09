---
title: Name.Creator Property (Excel)
keywords: vbaxl10.chm489074
f1_keywords:
- vbaxl10.chm489074
ms.prod: excel
api_name:
- Excel.Name.Creator
ms.assetid: 90c6fe07-e941-269f-71bf-e9dc6a982629
ms.date: 06/08/2017
---


# Name.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Name** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Name Object](name-object-excel.md)

