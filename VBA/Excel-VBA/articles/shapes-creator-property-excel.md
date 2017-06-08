---
title: Shapes.Creator Property (Excel)
keywords: vbaxl10.chm637074
f1_keywords:
- vbaxl10.chm637074
ms.prod: excel
api_name:
- Excel.Shapes.Creator
ms.assetid: 937cc87a-96a7-d1dc-7c06-0693f50293ea
ms.date: 06/08/2017
---


# Shapes.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Shapes** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Shapes Object](shapes-object-excel.md)

