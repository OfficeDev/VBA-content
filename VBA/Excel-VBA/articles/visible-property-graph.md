---
title: Visible Property (Graph)
keywords: vbagr10.chm66094
f1_keywords:
- vbagr10.chm66094
ms.prod: excel
ms.assetid: 8a2b1b7a-b880-0e43-ca9f-c5d2207f7cfd
ms.date: 06/08/2017
---


# Visible Property (Graph)

Visible property as it applies to the  **Application** object.

Determines whether the object is visible. Read/write Boolean.

 _expression_. **Visible**

 _expression_ Required. An expression that returns an **Application** object.
Visible property as it applies to the  **ChartFillFormat** object.
Determines whether the application is visible. Read/write MsoTriState .


|MsoTriState can be one of these MsoTriState constants.|
| **msoCTrue**|
| **msoFalse**|
| **msoTriStateMixed**|
| **msoTriStateToggle**|
| **msoTrue** The object is visible.|
 _expression_. **Visible**
 _expression_ Required. An expression that returns a **ChartFillFormat** object.

## Example

As it applies to the  **ChartFillFormat** object.

This example formats the chart's fill with a preset gradient and then makes the fill visible.




```vb
With myChart.ChartArea.Fill 
 .Visible = msoTrue 
 .PresetGradient msoGradientDiagonalDown, _ 
 3, msoGradientBrass 
End With
```


