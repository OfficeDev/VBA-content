---
title: Shading.ForegroundPatternColor Property (Word)
keywords: vbawd10.chm154796036
f1_keywords:
- vbawd10.chm154796036
ms.prod: word
api_name:
- Word.Shading.ForegroundPatternColor
ms.assetid: 2d8337e1-df14-8397-a59f-742fd03b0c4f
ms.date: 06/08/2017
---


# Shading.ForegroundPatternColor Property (Word)

Returns or sets the 24-bit color that's applied to the foreground of the  **Shading** object. This color is applied to the dots and lines in the shading pattern. Read/write.


## Syntax

 _expression_ . **ForegroundPatternColor**

 _expression_ Required. A variable that represents a **[Shading](shading-object-word.md)** object.


## Remarks

This property can be any valid  **WdColor** constant or a value returned by Visual Basic's **RGB** function.


## Example

This example applies shading with teal dots on a dark red background to the selection.


```vb
With Selection.Shading 
 .Texture = wdTexture30Percent 
 .ForegroundPatternColor = wdColorTeal 
 .BackgroundPatternColor = wdColorDarkRed 
End With
```


## See also


#### Concepts


[Shading Object](shading-object-word.md)

