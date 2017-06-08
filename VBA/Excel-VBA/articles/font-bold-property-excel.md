---
title: Font.Bold Property (Excel)
keywords: vbaxl10.chm559074
f1_keywords:
- vbaxl10.chm559074
ms.prod: excel
api_name:
- Excel.Font.Bold
ms.assetid: 7343989f-f973-0b1d-e595-c625ef2e0c15
ms.date: 06/08/2017
---


# Font.Bold Property (Excel)

 **True** if the font is bold. Read/write **Variant** .


## Syntax

 _expression_ . **Bold**

 _expression_ A variable that represents a **Font** object.


## Example

This example sets the font to bold for the range A1:A5 on Sheet1.


```vb
Worksheets("Sheet1").Range("A1:A5").Font.Bold = True
```


## See also


#### Concepts


[Font Object](font-object-excel.md)

