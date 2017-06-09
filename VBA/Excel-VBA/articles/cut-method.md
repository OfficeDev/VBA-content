---
title: Cut Method
keywords: vbagr10.chm5207276
f1_keywords:
- vbagr10.chm5207276
ms.prod: excel
api_name:
- Excel.Cut
ms.assetid: a0e35a76-9789-b661-e12b-04f11db84e3c
ms.date: 06/08/2017
---


# Cut Method

Cuts the specified range to the Clipboard or pastes it into a specified destination.

 _expression_. **Cut( _Destination_)**

 _expression_ Required. An expression that returns a **Range** object.

 **Destination** Optional **Variant**. The range where the object should be pasted. If this argument is omitted, the object is cut to the Clipboard.

## Example

This example cuts the range A1:G37 on the datasheet and places it on the Clipboard.


```vb
Set mySheet = myChart.Application.DataSheet 
mySheet.Range("A1:G37").Cut
```


