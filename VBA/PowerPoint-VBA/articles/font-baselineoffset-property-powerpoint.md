---
title: Font.BaselineOffset Property (PowerPoint)
keywords: vbapp10.chm575011
f1_keywords:
- vbapp10.chm575011
ms.prod: powerpoint
api_name:
- PowerPoint.Font.BaselineOffset
ms.assetid: aa948e2e-957c-ff4c-16b9-480d7f5f2d24
ms.date: 06/08/2017
---


# Font.BaselineOffset Property (PowerPoint)

Returns or sets the baseline offset for the specified superscript or subscript characters. Read/write.


## Syntax

 _expression_. **BaselineOffset**

 _expression_ A variable that represents a **Font** object.


### Return Value

Single


## Remarks

The  **BaselineOffset** property value be a floating-point value from - 1 through 1. A value of - 1 represents an offset of - 100 percent, and a value of 1 represents an offset of 100 percent.

Setting the  **BaselineOffset** property to a negative value automatically sets the **Subscript** property to **True** and the **Superscript** property to **False**.

Setting the  **BaselineOffset** property to a positive value automatically sets the **Subscript** property to **False** and the **Superscript** property to **True**.

Setting the  **Subscript** property to **True** automatically sets the **BaselineOffset** property to 0.3 (30 percent).

Setting the  **Superscript** property to **True** automatically sets the **BaselineOffset** property to - 0.25 ( - 25 percent).


## Example

This example sets the text for shape two on slide one and then makes the second character subscript with a 20-percent offset.


```vb
With Application.ActivePresentation.Slides(1) _
        .Shapes(2).TextFrame.TextRange
    .Text = "H2O"
    .Characters(2, 1).Font.BaselineOffset = -0.2
End With
```


## See also


#### Concepts


[Font Object](font-object-powerpoint.md)

