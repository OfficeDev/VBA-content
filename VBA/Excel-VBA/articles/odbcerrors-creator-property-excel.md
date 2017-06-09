---
title: ODBCErrors.Creator Property (Excel)
keywords: vbaxl10.chm528074
f1_keywords:
- vbaxl10.chm528074
ms.prod: excel
api_name:
- Excel.ODBCErrors.Creator
ms.assetid: 0db4a69d-36bd-a3cc-a407-e2a65bcf7fb3
ms.date: 06/08/2017
---


# ODBCErrors.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents an **ODBCErrors** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[ODBCErrors Object](odbcerrors-object-excel.md)

