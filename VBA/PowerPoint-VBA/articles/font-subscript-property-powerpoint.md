---
title: Font.Subscript Property (PowerPoint)
keywords: vbapp10.chm575009
f1_keywords:
- vbapp10.chm575009
ms.prod: powerpoint
api_name:
- PowerPoint.Font.Subscript
ms.assetid: ad23433b-b14b-9b2a-3bf6-772de41995f7
ms.date: 06/08/2017
---


# Font.Subscript Property (PowerPoint)

Determines whether the specified text is subscript. Read/write.


## Syntax

 _expression_. **Subscript**

 _expression_ A variable that represents a **Font** object.


### Return Value

MsoTriState


## Remarks

Setting the  **BaselineOffset** property to a negative value automatically sets the **Subscript** property to **msoTrue** and the **Superscript** property to **msoFalse**.

Setting the  **BaselineOffset** property to a positive value automatically sets the **Subscript** property to **msoFalse** and the **Superscript** property to **msoTrue**.

Setting the  **Subscript** property to **msoTrue** automatically sets the **BaselineOffset** property to ? 0.25 ( ? 25 percent).

The value of the  **Subscript** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified text is not subscript. The default.|
|**msoTriStateMixed**|Some characters are subscript and some aren't.|
|**msoTrue**|The specified text is subscript.|

## Example

This example enlarges the first character in the title on slide one if that character is subscript.


```vb
With Application.ActivePresentation.Slides(1) _
        .Shapes.Title.TextFrame.TextRange
    With .Characters(1, 1).Font
        If .Subscript Then
            scaleChar = -20 * .BaselineOffset
            .Size = .Size * scaleChar
        End If
    End With
End With
```


## See also


#### Concepts


[Font Object](font-object-powerpoint.md)

