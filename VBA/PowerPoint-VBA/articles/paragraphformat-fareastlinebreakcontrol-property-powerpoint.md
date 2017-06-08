---
title: ParagraphFormat.FarEastLineBreakControl Property (PowerPoint)
keywords: vbapp10.chm576012
f1_keywords:
- vbapp10.chm576012
ms.prod: powerpoint
api_name:
- PowerPoint.ParagraphFormat.FarEastLineBreakControl
ms.assetid: ffc0cb13-b547-5a33-e661-8a2cc4237e88
ms.date: 06/08/2017
---


# ParagraphFormat.FarEastLineBreakControl Property (PowerPoint)

Returns or sets the line break control option if you have an Asian language setting specified. Read/write.


## Syntax

 _expression_. **FarEastLineBreakControl**

 _expression_ A variable that represents a **ParagraphFormat** object.


### Return Value

MsoTriState


## Remarks

The value of the  **FarEastLineBreakControl** property can be one of these **MsoTriState** constants.


|||
|:-----|:-----|
|**msoFalse**|The line break control option is not selected.|
|**msoTrue**|The line break control option is selected.|

## Example

This example selects the line break option for the text in shape one on the first slide of the active presentation.


```vb
ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange.ParagraphFormat.FarEastLineBreakControl = msoTrue
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-powerpoint.md)

