---
title: CellFormat.Creator Property (Excel)
keywords: vbaxl10.chm675074
f1_keywords:
- vbaxl10.chm675074
ms.prod: excel
api_name:
- Excel.CellFormat.Creator
ms.assetid: 9a0b4160-9779-35dc-32bc-f750b490357d
ms.date: 06/08/2017
---


# CellFormat.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **CellFormat** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[CellFormat Object](cellformat-object-excel.md)

