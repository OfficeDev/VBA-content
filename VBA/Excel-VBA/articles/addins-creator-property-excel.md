---
title: AddIns.Creator Property (Excel)
keywords: vbaxl10.chm186074
f1_keywords:
- vbaxl10.chm186074
ms.prod: excel
api_name:
- Excel.AddIns.Creator
ms.assetid: 8fc7772e-1837-5336-9ae7-eca7f0dc14af
ms.date: 06/08/2017
---


# AddIns.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ An expression that returns a **AddIns** object.


### Return Value

XlCreator


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[AddIns Collection](addins-object-excel.md)

