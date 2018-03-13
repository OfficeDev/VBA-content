---
title: FillFormat.Patterned Method (PowerPoint)
keywords: vbapp10.chm552004
f1_keywords:
- vbapp10.chm552004
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.Patterned
ms.assetid: 665c5b1d-e2a2-64ab-a0c3-7d22d8d3121a
ms.date: 06/08/2017
---


# FillFormat.Patterned Method (PowerPoint)

Sets the specified fill to a pattern.


## Syntax

 _expression_. **Patterned**( **_Pattern_** )

 _expression_ A variable that represents a **FillFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Pattern_|Required|**MsoPatternType**|The pattern to be used for the specified fill. See Reamrks for possible values.|

## Remarks

Use the [BackColor](fillformat-backcolor-property-powerpoint.md)and  **[ForeColor](fillformat-forecolor-property-powerpoint.md)** properties to set the colors used in the pattern.

The value of the Pattern parameter can be one of these  **MsoPatternType** constants.


||
|:-----|
|<strong>msoPattern10Percent</strong>|
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
|
<strong>msoPatternDarkVertical</strong>|

## Example

This example adds an oval with a patterned fill to  `myDocument`.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddShape(msoShapeOval, 60, 60, 80, 40).Fill

    .ForeColor.RGB = RGB(128, 0, 0)

    .BackColor.RGB = RGB(0, 0, 255)

    .Patterned msoPatternDarkVertical

End With
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-powerpoint.md)

