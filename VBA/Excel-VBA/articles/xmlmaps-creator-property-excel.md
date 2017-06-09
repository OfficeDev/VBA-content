---
title: XmlMaps.Creator Property (Excel)
keywords: vbaxl10.chm755074
f1_keywords:
- vbaxl10.chm755074
ms.prod: excel
api_name:
- Excel.XmlMaps.Creator
ms.assetid: f0973a6d-dc2a-11ae-3c52-d73807098d62
ms.date: 06/08/2017
---


# XmlMaps.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **XmlMaps** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[XmlMaps Object](xmlmaps-object-excel.md)

