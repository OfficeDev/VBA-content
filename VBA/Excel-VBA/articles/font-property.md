---
title: Font Property
keywords: vbagr10.chm65682
f1_keywords:
- vbagr10.chm65682
ms.prod: excel
api_name:
- Excel.Font
ms.assetid: 0bc46ec4-998e-043e-0713-9a381ec2b6ad
ms.date: 06/08/2017
---


# Font Property

Returns a Font object that represents the font of the specified object. Read/write Font object only for the DataSheet object, for all other objects, read-only Font object.

 _expression_. **Font**

 _expression_ Required. An expression that returns one of the above objects.


## Example

This example sets the font in the chart title to 14-point bold italic.


```vb
With myChart.ChartTitle.Font 
 .Size = 14 
 .Bold = True 
 .Italic = True 
End With 

```


