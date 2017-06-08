---
title: ListObject.Creator Property (Excel)
keywords: vbaxl10.chm733074
f1_keywords:
- vbaxl10.chm733074
ms.prod: excel
api_name:
- Excel.ListObject.Creator
ms.assetid: 39d04a9a-c36e-5d09-df79-cbb802ddbe28
ms.date: 06/08/2017
---


# ListObject.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **ListObject** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[ListObject Object](listobject-object-excel.md)

