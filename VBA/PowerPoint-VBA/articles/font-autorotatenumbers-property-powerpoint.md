---
title: Font.AutoRotateNumbers Property (PowerPoint)
keywords: vbapp10.chm575018
f1_keywords:
- vbapp10.chm575018
ms.prod: powerpoint
api_name:
- PowerPoint.Font.AutoRotateNumbers
ms.assetid: 621ccc86-d5cb-d2c1-262f-5652eff5800a
ms.date: 06/08/2017
---


# Font.AutoRotateNumbers Property (PowerPoint)

Returns or sets lateral compression. Read/write.


## Syntax

 _expression_. **AutoRotateNumbers**

 _expression_ A variable that represents an **Font** object.


### Return Value

MsoTriState


## Remarks

The value of the  **AutoRotateNumbers** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**| Half-width numbers are not compressed in lateral columns.|
|**msoTrue**|Displays half-width numbers within vertical text in two-character lateral columns.|

## Example

This example sets the text direction of shape three on the first slide to vertical text, and sets lateral column compression.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3).TextFrame

    .Orientation = msoTextOrientationVerticalFarEast

    .TextRange.Font.AutoRotateNumbers = msoTrue

End With
```


## See also


#### Concepts


[Font Object](font-object-powerpoint.md)

