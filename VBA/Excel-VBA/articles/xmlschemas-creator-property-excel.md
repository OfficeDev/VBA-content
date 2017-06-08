---
title: XmlSchemas.Creator Property (Excel)
keywords: vbaxl10.chm751074
f1_keywords:
- vbaxl10.chm751074
ms.prod: excel
api_name:
- Excel.XmlSchemas.Creator
ms.assetid: c9000e23-0426-9571-8104-6b4542f661fa
ms.date: 06/08/2017
---


# XmlSchemas.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **XmlSchemas** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[XmlSchemas Object](xmlschemas-object-excel.md)

