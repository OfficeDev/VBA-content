---
title: Workbooks.Creator Property (Excel)
keywords: vbaxl10.chm202074
f1_keywords:
- vbaxl10.chm202074
ms.prod: excel
api_name:
- Excel.Workbooks.Creator
ms.assetid: 26f90d17-3dc1-3c35-3ddb-ddcdf4e99998
ms.date: 06/08/2017
---


# Workbooks.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Workbooks** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Workbooks Object](workbooks-object-excel.md)

