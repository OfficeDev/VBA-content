---
title: Interior.PatternColorIndex Property (Word)
keywords: vbawd10.chm2818058
f1_keywords:
- vbawd10.chm2818058
ms.prod: word
api_name:
- Word.Interior.PatternColorIndex
ms.assetid: 2f2400e1-1995-2996-01a3-fd5ff0e6bf47
ms.date: 06/08/2017
---


# Interior.PatternColorIndex Property (Word)

Returns or sets the color of the interior pattern as an index into the current color palette, or as one of the following  **[XlColorIndex](xlcolorindex-enumeration-word.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone** . Read/write **Long** .


## Syntax

 _expression_ . **PatternColorIndex**

 _expression_ A variable that represents an **[Interior](interior-object-word.md)** object.


## Remarks

Set this property to  **xlColorIndexAutomatic** to specify the automatic fill style for drawing objects. Set this property to **xlColorIndexNone** to specify that you do not want a pattern (this is the same as setting the **Pattern** property of the **Interior** object to **xlPatternNone** ).


## Example

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


[Interior Object](interior-object-word.md)

