---
title: Pattern Property
keywords: vbagr10.chm65631
f1_keywords:
- vbagr10.chm65631
ms.prod: excel
api_name:
- Excel.Pattern
ms.assetid: 3cc8475d-dc65-b2eb-e1ba-2bd95c5c0b03
ms.date: 06/08/2017
---


# Pattern Property

For the ChartFillFormat object, returns or sets the fill pattern, read-only MsoPatternType . For the Interior object, returns or sets the interior pattern, read/write Variant.



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

 _expression_. **Pattern**

 _expression_ Required. An expression that returns one of the above objects.

## Example

This example adds a crisscross pattern to the interior of the plot area.


```
myChart.PlotArea.Interior.Pattern = xlPatternCrissCross
```


