---
title: ParagraphFormat.HangingPunctuation Property (PowerPoint)
keywords: vbapp10.chm576014
f1_keywords:
- vbapp10.chm576014
ms.prod: powerpoint
api_name:
- PowerPoint.ParagraphFormat.HangingPunctuation
ms.assetid: e7e1f5b2-e0ed-9b5c-7c14-fcf4c134e3bb
ms.date: 06/08/2017
---


# ParagraphFormat.HangingPunctuation Property (PowerPoint)

Returns or sets the hanging punctuation option if you have an Asian language setting specified. Read/write.


## Syntax

 _expression_. **HangingPunctuation**

 _expression_ A variable that represents a **ParagraphFormat** object.


### Return Value

MsoTriState


## Remarks

The value of the  **HangingPunctuation** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The hanging punctuation option is not selected.|
|**msoTrue**| The hanging punctuation option is selected.|

## Example

This example selects hanging punctuation for the first paragraph of the active presentation.


```vb
ActivePresentation.Paragraphs(1).HangingPunctuation = msoTrue
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-powerpoint.md)

