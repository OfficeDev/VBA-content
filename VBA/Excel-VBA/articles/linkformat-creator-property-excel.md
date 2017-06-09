---
title: LinkFormat.Creator Property (Excel)
keywords: vbaxl10.chm633074
f1_keywords:
- vbaxl10.chm633074
ms.prod: excel
api_name:
- Excel.LinkFormat.Creator
ms.assetid: cb1b0a6d-af14-0f9c-2e5e-d991d7011a20
ms.date: 06/08/2017
---


# LinkFormat.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **LinkFormat** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[LinkFormat Object](linkformat-object-excel.md)

