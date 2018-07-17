---
title: ParagraphFormat.TextDirection Property (PowerPoint)
keywords: vbapp10.chm576015
f1_keywords:
- vbapp10.chm576015
ms.prod: powerpoint
api_name:
- PowerPoint.ParagraphFormat.TextDirection
ms.assetid: 42b8cd29-c467-07c9-c9c9-f644fdc824ae
ms.date: 06/08/2017
---


# ParagraphFormat.TextDirection Property (PowerPoint)

Returns or sets the text direction for the specified paragraph. Read/write.


## Syntax

 _expression_. **TextDirection**

 _expression_ A variable that represents a **ParagraphFormat** object.


### Return Value

PpDirection


## Remarks

The default value depends on the language support you have selected or installed.

The value of the  **TextDirection** property can be one of these **PpDirection** constants.


||
|:-----|
|**ppDirectionLeftToRight**|
|**ppDirectionMixed**|
|**ppDirectionRightToLeft**|

## Example

This example displays the text direction for the paragraphs in shape two on slide one in the active presentation.


```vb
MsgBox ActivePresentation.Slides(1).Shapes(2).TextFrame.TextRange _
    .ParagraphFormat.TextDirection
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-powerpoint.md)

