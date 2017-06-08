---
title: ListDataFormat.Creator Property (Excel)
keywords: vbaxl10.chm757074
f1_keywords:
- vbaxl10.chm757074
ms.prod: excel
api_name:
- Excel.ListDataFormat.Creator
ms.assetid: f8ac98f1-f34a-430c-16fa-d62d07c76276
ms.date: 06/08/2017
---


# ListDataFormat.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **ListDataFormat** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[ListDataFormat Object](listdataformat-object-excel.md)

