---
title: PublishObjects.Creator Property (Excel)
keywords: vbaxl10.chm649074
f1_keywords:
- vbaxl10.chm649074
ms.prod: excel
api_name:
- Excel.PublishObjects.Creator
ms.assetid: 10cbdf25-3e7e-4bc5-8a32-8dbe2f7bbb46
ms.date: 06/08/2017
---


# PublishObjects.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **PublishObjects** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[PublishObjects Object](publishobjects-object-excel.md)

