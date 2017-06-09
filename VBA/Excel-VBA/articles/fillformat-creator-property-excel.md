---
title: FillFormat.Creator Property (Excel)
ms.prod: excel
api_name:
- Excel.FillFormat.Creator
ms.assetid: f4e02d6c-49b7-d837-c090-096975d8efb1
ms.date: 06/08/2017
---


# FillFormat.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **FillFormat** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[FillFormat Object](fillformat-object-excel.md)

