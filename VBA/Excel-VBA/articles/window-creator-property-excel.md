---
title: Window.Creator Property (Excel)
keywords: vbaxl10.chm355074
f1_keywords:
- vbaxl10.chm355074
ms.prod: excel
api_name:
- Excel.Window.Creator
ms.assetid: fb41f6ad-241a-3a04-729f-f04e1c5d0296
ms.date: 06/08/2017
---


# Window.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Window** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Window Object](window-object-excel.md)

