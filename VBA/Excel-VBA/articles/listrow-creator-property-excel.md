---
title: ListRow.Creator Property (Excel)
keywords: vbaxl10.chm741074
f1_keywords:
- vbaxl10.chm741074
ms.prod: excel
api_name:
- Excel.ListRow.Creator
ms.assetid: 3b750487-3ea6-815b-0389-55313cb2f36b
ms.date: 06/08/2017
---


# ListRow.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **ListRow** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[ListRow Object](listrow-object-excel.md)

