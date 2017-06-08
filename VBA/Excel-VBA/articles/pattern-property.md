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

 _expression_. **Pattern**

 _expression_ Required. An expression that returns one of the above objects.

## Example

This example adds a crisscross pattern to the interior of the plot area.


```
myChart.PlotArea.Interior.Pattern = xlPatternCrissCross
```


