---
title: Style.Creator Property (Excel)
keywords: vbaxl10.chm176074
f1_keywords:
- vbaxl10.chm176074
ms.prod: excel
api_name:
- Excel.Style.Creator
ms.assetid: d7473e53-fba0-a195-7dba-430e3b6d1df6
ms.date: 06/08/2017
---


# Style.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Style** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Style Object](style-object-excel.md)

