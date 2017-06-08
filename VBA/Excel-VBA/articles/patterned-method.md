---
title: Patterned Method
keywords: vbagr10.chm67164
f1_keywords:
- vbagr10.chm67164
ms.prod: excel
api_name:
- Excel.Patterned
ms.assetid: a492f089-cd6e-e7c3-2b25-7bcfadde4319
ms.date: 06/08/2017
---


# Patterned Method

Sets a pattern for the specified fill.

 _expression_. **Patterned**( **_Pattern_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **Pattern**Required 
 **MsoPatternType**
. The type of pattern.


|MsoPatternType can be one of these MsoPatternType constants.|
| **msoPattern10Percent**|
| **msoPattern20Percent**|
| **msoPattern25Percent**|
| **msoPattern30Percent**|
| **msoPattern40Percent**|
| **msoPattern50Percent**|
| **msoPattern5Percent**|
| **msoPattern60Percent**|
| **msoPattern70Percent**|
| **msoPattern75Percent**|
| **msoPattern80Percent**|
| **msoPattern90Percent**|
| **msoPatternDarkDownwardDiagonal**|
| **msoPatternDarkHorizontal**|
| **msoPatternDarkUpwardDiagonal**|
| **msoPatternDarkVertical**|
| **msoPatternDashedDownwardDiagonal**|
| **msoPatternDashedHorizontal**|
| **msoPatternDashedUpwardDiagonal**|
| **msoPatternDashedVertical**|
| **msoPatternDiagonalBrick**|
| **msoPatternDivot**|
| **msoPatternDottedDiamond**|
| **msoPatternDottedGrid**|
| **msoPatternHorizontalBrick**|
| **msoPatternLargeCheckerBoard**|
| **msoPatternLargeConfetti**|
| **msoPatternLargeGrid**|
| **msoPatternLightDownwardDiagonal**|
| **msoPatternLightHorizontal**|
| **msoPatternLightUpwardDiagonal**|
| **msoPatternLightVertical**|
| **msoPatternMixed**|
| **msoPatternNarrowHorizontal**|
| **msoPatternNarrowVertical**|
| **msoPatternOutlinedDiamond**|
| **msoPatternPlaid**|
| **msoPatternShingle**|
| **msoPatternSmallCheckerBoard**|
| **msoPatternSmallConfetti**|
| **msoPatternSmallGrid**|
| **msoPatternSolidDiamond**|
| **msoPatternSphere**|
| **msoPatternTrellis**|
| **msoPatternWave**|
| **msoPatternWeave**|
| **msoPatternWideDownwardDiagonal**|
| **msoPatternWideUpwardDiagonal**|
| **msoPatternZigZag**|

## Example

This example sets the fill pattern.


```vb
With myChart.ChartArea.Fill 
 .Patterned msoPatternDiagonalBrick 
 .Visible = True 
End With
```


