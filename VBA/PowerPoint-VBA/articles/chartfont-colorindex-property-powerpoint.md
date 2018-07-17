---
title: ChartFont.ColorIndex Property (PowerPoint)
keywords: vbapp10.chm704004
f1_keywords:
- vbapp10.chm704004
ms.prod: powerpoint
api_name:
- PowerPoint.ChartFont.ColorIndex
ms.assetid: 2f0765bf-9a3b-999a-2dd6-17009fbd619d
ms.date: 06/08/2017
---


# ChartFont.ColorIndex Property (PowerPoint)

Returns or sets the color of the font. Read/write  **Variant**.


## Syntax

 _expression_. **ColorIndex**

 _expression_ A variable that represents a **[ChartFont](chartfont-object-powerpoint.md)** object.


## Remarks

The color is specified as an index value into the current color palette, or as one of the following  **[XlColorIndex](xlcolorindex-enumeration-powerpoint.md)** constants:


-  **xlColorIndexAutomatic**
    
-  **xlColorIndexNone**
    

## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example changes the font color in the title of the first chart in the active document to red.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Title

                ' Set the color to red.

                .Font.ColorIndex = 3

            End If

        End With

    End If

End With
```


## See also


#### Concepts


[ChartFont Object](chartfont-object-powerpoint.md)

