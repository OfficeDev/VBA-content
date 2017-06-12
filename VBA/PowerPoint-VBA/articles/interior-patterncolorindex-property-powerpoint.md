---
title: Interior.PatternColorIndex Property (PowerPoint)
keywords: vbapp10.chm707006
f1_keywords:
- vbapp10.chm707006
ms.prod: powerpoint
api_name:
- PowerPoint.Interior.PatternColorIndex
ms.assetid: d7a42e0c-d3f4-85a1-009c-0b6d2385ee77
ms.date: 06/08/2017
---


# Interior.PatternColorIndex Property (PowerPoint)

Returns or sets the color of the interior pattern as an index into the current color palette, or as one of the following  **[XlColorIndex](xlcolorindex-enumeration-powerpoint.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone**. Read/write **Long**.


## Syntax

 _expression_. **PatternColorIndex**

 _expression_ A variable that represents an **[Interior](interior-object-powerpoint.md)** object.


## Remarks

Set this property to  **xlColorIndexAutomatic** to specify the automatic fill style for drawing objects. Set this property to **xlColorIndexNone** to specify that you do not want a pattern (this is the same as setting the **Pattern** property of the **Interior** object to **xlPatternNone** ).


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables up and down bars, then adds a criss-cross pattern to the down bars and sets the pattern color to red, for the first chart group of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.ChartGroups(1)

            .HasUpDownBars = True

            .DownBars.Interior.Pattern = xlPatternCrissCross

            .DownBars.Interior.PatternColorIndex = 3

        End With

    End If

End With
```


## See also


#### Concepts


[Interior Object](interior-object-powerpoint.md)

