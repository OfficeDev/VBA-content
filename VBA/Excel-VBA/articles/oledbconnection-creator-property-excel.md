---
title: OLEDBConnection.Creator Property (Excel)
keywords: vbaxl10.chm793074
f1_keywords:
- vbaxl10.chm793074
ms.prod: excel
api_name:
- Excel.OLEDBConnection.Creator
ms.assetid: a2a5b5cd-9fea-0756-d2a6-ff632a29ffa9
ms.date: 06/08/2017
---


# OLEDBConnection.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents an **OLEDBConnection** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[OLEDBConnection Object](oledbconnection-object-excel.md)

