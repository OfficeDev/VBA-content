---
title: ODBCError.Creator Property (Excel)
keywords: vbaxl10.chm526074
f1_keywords:
- vbaxl10.chm526074
ms.prod: excel
api_name:
- Excel.ODBCError.Creator
ms.assetid: 0c565d02-2e5e-e997-f3ea-0775121eb545
ms.date: 06/08/2017
---


# ODBCError.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents an **ODBCError** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[ODBCError Object](odbcerror-object-excel.md)

