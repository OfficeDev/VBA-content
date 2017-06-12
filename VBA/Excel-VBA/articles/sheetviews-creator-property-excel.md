---
title: SheetViews.Creator Property (Excel)
keywords: vbaxl10.chm791074
f1_keywords:
- vbaxl10.chm791074
ms.prod: excel
api_name:
- Excel.SheetViews.Creator
ms.assetid: 21ab3bb7-6269-db13-d81d-eda3861aa846
ms.date: 06/08/2017
---


# SheetViews.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **SheetViews** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[SheetViews Object](sheetviews-object-excel.md)

