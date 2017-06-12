---
title: Font.Superscript Property (PowerPoint)
keywords: vbapp10.chm575010
f1_keywords:
- vbapp10.chm575010
ms.prod: powerpoint
api_name:
- PowerPoint.Font.Superscript
ms.assetid: 6f0bba73-f375-d715-3ddb-f1ab6041336c
ms.date: 06/08/2017
---


# Font.Superscript Property (PowerPoint)

Determines whether the specified text is superscript. Read/write.


## Syntax

 _expression_. **Superscript**

 _expression_ A variable that represents a **Font** object.


### Return Value

MsoTriState


## Remarks

Setting the  **BaselineOffset** property to a negative value automatically sets the **Subscript** property to **msoTrue** and the **Superscript** property to **msoFalse**.

Setting the  **BaselineOffset** property to a positive value automatically sets the **Subscript** property to **msoFalse** and the **Superscript** property to **msoTrue**.

Setting the  **Superscript** property to **msoTrue** automatically sets the **BaselineOffset** property to 0.3 (30 percent).

The value of the  **Superscript** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified text is not superscript. The default.|
|**msoTriStateMixed**|Some characters are superscript and some aren't.|
|**msoTrue**|The specified text is superscript.|

## Example

This example sets the text for shape two on slide one and then makes the fifth character superscript with a 30-percent offset.


```vb
With Application.ActivePresentation.Slides(1).Shapes(2).TextFrame

    With .TextRange

        .Text = "E=mc2"

        .Characters(5, 1).Font.Superscript = msoTrue

    End With

End With
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-powerpoint.md)
[Font Object](font-object-powerpoint.md)

