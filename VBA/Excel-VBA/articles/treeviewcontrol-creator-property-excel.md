---
title: TreeviewControl.Creator Property (Excel)
keywords: vbaxl10.chm665074
f1_keywords:
- vbaxl10.chm665074
ms.prod: excel
api_name:
- Excel.TreeviewControl.Creator
ms.assetid: b8956992-0bc3-f98a-1155-5c1f3a0f3ec6
ms.date: 06/08/2017
---


# TreeviewControl.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **TreeviewControl** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[TreeviewControl Object](treeviewcontrol-object-excel.md)

