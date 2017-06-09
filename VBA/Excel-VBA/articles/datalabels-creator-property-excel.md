---
title: DataLabels.Creator Property (Excel)
keywords: vbaxl10.chm583074
f1_keywords:
- vbaxl10.chm583074
ms.prod: excel
api_name:
- Excel.DataLabels.Creator
ms.assetid: 86b4beaa-9a5f-ca3e-7e4f-78b905e94bae
ms.date: 06/08/2017
---


# DataLabels.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **DataLabels** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[DataLabels Object](datalabels-object-excel.md)

