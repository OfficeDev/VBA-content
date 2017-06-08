---
title: TickLabels.Creator Property (Excel)
keywords: vbaxl10.chm615074
f1_keywords:
- vbaxl10.chm615074
ms.prod: excel
api_name:
- Excel.TickLabels.Creator
ms.assetid: e319d5b6-9048-8823-0161-cba3ed4037f4
ms.date: 06/08/2017
---


# TickLabels.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **TickLabels** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[TickLabels Object](ticklabels-object-excel.md)

