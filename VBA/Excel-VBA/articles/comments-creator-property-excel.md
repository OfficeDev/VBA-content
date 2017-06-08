---
title: Comments.Creator Property (Excel)
keywords: vbaxl10.chm513074
f1_keywords:
- vbaxl10.chm513074
ms.prod: excel
api_name:
- Excel.Comments.Creator
ms.assetid: 839fc2bb-e9d8-a998-803e-169100dc8ef2
ms.date: 06/08/2017
---


# Comments.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Comments** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Comments Object](comments-object-excel.md)

