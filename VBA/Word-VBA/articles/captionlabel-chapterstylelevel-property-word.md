---
title: CaptionLabel.ChapterStyleLevel Property (Word)
keywords: vbawd10.chm158924805
f1_keywords:
- vbawd10.chm158924805
ms.prod: word
api_name:
- Word.CaptionLabel.ChapterStyleLevel
ms.assetid: c0824b64-8709-009a-53cd-353238289e88
ms.date: 06/08/2017
---


# CaptionLabel.ChapterStyleLevel Property (Word)

Returns or sets the heading style that marks a new chapter when chapter numbers are included with the specified caption label. Read/write  **Long** .


## Syntax

 _expression_ . **ChapterStyleLevel**

 _expression_ A variable that represents a **[CaptionLabel](captionlabel-object-word.md)** object.


## Remarks

The number 1 corresponds to Heading 1, 2 corresponds to Heading 2, and so on. The  **[IncludeChapterNumber](captionlabel-includechapternumber-property-word.md)** property must be set to **True** for chapter numbers to be included with caption labels.


## Example

This example formats the table's caption label to include a chapter number. The chapter number is taken from paragraphs formatted with the Heading 2 style.


```vb
With CaptionLabels(wdCaptionTable) 
 .IncludeChapterNumber = True 
 .ChapterStyleLevel = 2 
End With
```


## See also


#### Concepts


[CaptionLabel Object](captionlabel-object-word.md)

