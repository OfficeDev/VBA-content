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
| <strong>msoPattern10Percent</strong>|
| 
<strong>msoPattern20Percent</strong>|
| 
<strong>msoPattern25Percent</strong>|
| 
<strong>msoPattern30Percent</strong>|
| 
<strong>msoPattern40Percent</strong>|
| 
<strong>msoPattern50Percent</strong>|
| 
<strong>msoPattern5Percent</strong>|
| 
<strong>msoPattern60Percent</strong>|
| 
<strong>msoPattern70Percent</strong>|
| 
<strong>msoPattern75Percent</strong>|
| 
<strong>msoPattern80Percent</strong>|
| 
<strong>msoPattern90Percent</strong>|
| 
<strong>msoPatternDarkDownwardDiagonal</strong>|
| 
<strong>msoPatternDarkHorizontal</strong>|
| 
<strong>msoPatternDarkUpwardDiagonal</strong>|
| 
<strong>msoPatternDarkVertical</strong>|
| 
<strong>msoPatternDashedDownwardDiagonal</strong>|
| 
<strong>msoPatternDashedHorizontal</strong>|
| 
<strong>msoPatternDashedUpwardDiagonal</strong>|
| 
<strong>msoPatternDashedVertical</strong>|
| 
<strong>msoPatternDiagonalBrick</strong>|
| 
<strong>msoPatternDivot</strong>|
| 
<strong>msoPatternDottedDiamond</strong>|
| 
<strong>msoPatternDottedGrid</strong>|
| 
<strong>msoPatternHorizontalBrick</strong>|
| 
<strong>msoPatternLargeCheckerBoard</strong>|
| 
<strong>msoPatternLargeConfetti</strong>|
| 
<strong>msoPatternLargeGrid</strong>|
| 
<strong>msoPatternLightDownwardDiagonal</strong>|
| 
<strong>msoPatternLightHorizontal</strong>|
| 
<strong>msoPatternLightUpwardDiagonal</strong>|
| 
<strong>msoPatternLightVertical</strong>|
| 
<strong>msoPatternMixed</strong>|
| 
<strong>msoPatternNarrowHorizontal</strong>|
| 
<strong>msoPatternNarrowVertical</strong>|
| 
<strong>msoPatternOutlinedDiamond</strong>|
| 
<strong>msoPatternPlaid</strong>|
| 
<strong>msoPatternShingle</strong>|
| 
<strong>msoPatternSmallCheckerBoard</strong>|
| 
<strong>msoPatternSmallConfetti</strong>|
| 
<strong>msoPatternSmallGrid</strong>|
| 
<strong>msoPatternSolidDiamond</strong>|
| 
<strong>msoPatternSphere</strong>|
| 
<strong>msoPatternTrellis</strong>|
| 
<strong>msoPatternWave</strong>|
| 
<strong>msoPatternWeave</strong>|
| 
<strong>msoPatternWideDownwardDiagonal</strong>|
| 
<strong>msoPatternWideUpwardDiagonal</strong>|
| 
<strong>msoPatternZigZag</strong>|

## Example

This example sets the fill pattern.


```vb
With myChart.ChartArea.Fill 
 .Patterned msoPatternDiagonalBrick 
 .Visible = True 
End With
```


