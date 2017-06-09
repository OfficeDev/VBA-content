---
title: DialogSheetView.Creator Property (Excel)
keywords: vbaxl10.chm786074
f1_keywords:
- vbaxl10.chm786074
ms.prod: excel
api_name:
- Excel.DialogSheetView.Creator
ms.assetid: 7118a311-7f47-f229-78a5-6b1fec2d7fd9
ms.date: 06/08/2017
---


# DialogSheetView.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **DialogSheetView** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[DialogSheetView Object](dialogsheetview-object-excel.md)

