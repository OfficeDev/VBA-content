---
title: WorkbookConnection.Creator Property (Excel)
keywords: vbaxl10.chm773074
f1_keywords:
- vbaxl10.chm773074
ms.prod: excel
api_name:
- Excel.WorkbookConnection.Creator
ms.assetid: b8862979-d128-cd86-31ef-19515741792c
ms.date: 06/08/2017
---


# WorkbookConnection.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **WorkbookConnection** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[WorkbookConnection Object](workbookconnection-object-excel.md)

