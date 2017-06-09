---
title: Interior.PatternColor Property (PowerPoint)
keywords: vbapp10.chm707005
f1_keywords:
- vbapp10.chm707005
ms.prod: powerpoint
api_name:
- PowerPoint.Interior.PatternColor
ms.assetid: eb8e0993-fe73-3ab5-3b89-e5a306b20149
ms.date: 06/08/2017
---


# Interior.PatternColor Property (PowerPoint)

Returns or sets the color of the interior pattern as an RGB value. Read/write  **Variant**.


## Syntax

 _expression_. **PatternColor**

 _expression_ A variable that represents an **[Interior](interior-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables up and down bars, then adds a criss-cross pattern to the down bars and sets the pattern color to blue, for the first chart group of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.ChartGroups(1)

            .HasUpDownBars = True

            .DownBars.Interior.Pattern = xlPatternCrissCross

            .DownBars.Interior.PatternColor = RGB(0, 0, 255)

        End With

    End If

End With
```


## See also


#### Concepts


[Interior Object](interior-object-powerpoint.md)

