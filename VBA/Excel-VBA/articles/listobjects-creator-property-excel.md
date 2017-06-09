---
title: ListObjects.Creator Property (Excel)
keywords: vbaxl10.chm731074
f1_keywords:
- vbaxl10.chm731074
ms.prod: excel
api_name:
- Excel.ListObjects.Creator
ms.assetid: 6baa548b-04a6-e0eb-d45f-8d3f24848c3b
ms.date: 06/08/2017
---


# ListObjects.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **ListObjects** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[ListObjects Object](listobjects-object-excel.md)

