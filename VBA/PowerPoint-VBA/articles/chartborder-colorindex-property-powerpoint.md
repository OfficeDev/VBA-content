---
title: ChartBorder.ColorIndex Property (PowerPoint)
keywords: vbapp10.chm685002
f1_keywords:
- vbapp10.chm685002
ms.prod: powerpoint
api_name:
- PowerPoint.ChartBorder.ColorIndex
ms.assetid: c6601494-e82d-f67f-3cd9-bb7fa9c8ab3f
ms.date: 06/08/2017
---


# ChartBorder.ColorIndex Property (PowerPoint)

Returns or sets the color of the border. Read/write  **Variant**.


## Syntax

 _expression_. **ColorIndex**

 _expression_ A variable that represents a **[ChartBorder](chartborder-object-powerpoint.md)** object.


## Remarks

The color is specified as an index value into the current color palette, or as one of the following  **[XlColorIndex](xlcolorindex-enumeration-powerpoint.md)** constants:


-  **xlColorIndexAutomatic**
    
-  **xlColorIndexNone**
    

## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the color of the major gridlines for the value axis of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlValue)

            If .HasMajorGridlines Then

                ' Set the color to blue.

                .MajorGridlines.Border.ColorIndex = 5

            End If

        End With

    End If

End With
```


## See also


#### Concepts


[ChartBorder Object](chartborder-object-powerpoint.md)

