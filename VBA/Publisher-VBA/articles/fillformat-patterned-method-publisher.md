---
title: FillFormat.Patterned Method (Publisher)
keywords: vbapb10.chm2359314
f1_keywords:
- vbapb10.chm2359314
ms.prod: publisher
api_name:
- Publisher.FillFormat.Patterned
ms.assetid: 10e363b7-1160-55d3-5c97-733b7742b619
ms.date: 06/08/2017
---


# FillFormat.Patterned Method (Publisher)

Sets the specified fill to a pattern.


## Syntax

 _expression_. **Patterned**( **_Pattern_**)

 _expression_A variable that represents a  **FillFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Pattern|Required| **MsoPatternType**|The pattern to be used for the specified fill.|

## Remarks

The Pattern parameter can be one of the  **MsoPatternType** constants declared in the Microsoft Office type library and shown in the following table.



| **msoPattern5Percent**|
| **msoPattern10Percent**|
| **msoPattern20Percent**|
| **msoPattern25Percent**|
| **msoPattern30Percent**|
| **msoPattern40Percent**|
| **msoPattern50Percent**|
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
Use the  [BackColor](fillformat-backcolor-property-publisher.md)and  [ForeColor](fillformat-forecolor-property-publisher.md)properties to set the colors used in the pattern.


## Example

This example adds an oval with a patterned fill to the active publication.


```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeOval, _ 
 Left:=60, Top:=60, Width:=80, Height:=40).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .BackColor.RGB = RGB(0, 0, 255) 
 .Patterned Pattern:=msoPatternDarkVertical 
End With 

```


