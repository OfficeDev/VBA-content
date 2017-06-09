---
title: ParagraphFormat.BaseLineAlignment Property (PowerPoint)
keywords: vbapp10.chm576011
f1_keywords:
- vbapp10.chm576011
ms.prod: powerpoint
api_name:
- PowerPoint.ParagraphFormat.BaseLineAlignment
ms.assetid: b59f680f-a5a9-f6bc-85d5-f14670269ae8
ms.date: 06/08/2017
---


# ParagraphFormat.BaseLineAlignment Property (PowerPoint)

Returns or sets the base line alignment for the specified paragraph. Read/write.


## Syntax

 _expression_. **BaseLineAlignment**

 _expression_ A variable that represents a **ParagraphFormat** object.


### Return Value

PpBaselineAlignment


## Remarks

The value of the  **BaseLineAlignment** property can be one of these **PpBaselineAlignment** constants


||
|:-----|
|**ppBaselineAlignBaseline**|
|**ppBaselineAlignCenter**|
|**ppBaselineAlignFarEast50**|
|**ppBaselineAlignMixed**|
|**ppBaselineAlignTop**|

## Example

This example displays the base line alignment for the paragraphs in shape two on slide one in the active presentation.


```vb
MsgBox ActivePresentation.Slides(1).Shapes(2).TextFrame.TextRange _
    .ParagraphFormat.BaseLineAlignment
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-powerpoint.md)

