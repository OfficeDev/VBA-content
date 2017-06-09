---
title: Actions.Creator Property (Excel)
keywords: vbaxl10.chm799074
f1_keywords:
- vbaxl10.chm799074
ms.prod: excel
api_name:
- Excel.Actions.Creator
ms.assetid: 453bcc61-d054-47b8-3447-f2896c343699
ms.date: 06/08/2017
---


# Actions.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ An expression that returns a **Actions** object.


### Return Value

XlCreator


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Actions Object](actions-object-excel.md)

