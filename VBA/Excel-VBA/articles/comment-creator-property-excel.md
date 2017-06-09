---
title: Comment.Creator Property (Excel)
keywords: vbaxl10.chm515074
f1_keywords:
- vbaxl10.chm515074
ms.prod: excel
api_name:
- Excel.Comment.Creator
ms.assetid: c0765a60-a99b-ea10-b708-549222dc4e95
ms.date: 06/08/2017
---


# Comment.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Comment** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Comment Object](comment-object-excel.md)

