---
title: ChartBorder.ColorIndex Property (Word)
keywords: vbawd10.chm61014018
f1_keywords:
- vbawd10.chm61014018
ms.prod: word
api_name:
- Word.ChartBorder.ColorIndex
ms.assetid: e9457184-7100-9482-398e-cc7f11e4b05c
ms.date: 06/08/2017
---


# ChartBorder.ColorIndex Property (Word)

Returns or sets the color of the border. Read/write  **Variant** .


## Syntax

 _expression_ . **ColorIndex**

 _expression_ A variable that represents a **[ChartBorder](chartborder-object-word.md)** object.


## Remarks

The color is specified as an index value into the current color palette, or as one of the following  **[XlColorIndex](xlcolorindex-enumeration-word.md)** constants:


-  **xlColorIndexAutomatic**
    
-  **xlColorIndexNone**
    

## Example

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


[ChartBorder Object](chartborder-object-word.md)

