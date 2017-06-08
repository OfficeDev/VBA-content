---
title: CaptionLabel.IncludeChapterNumber Property (Word)
keywords: vbawd10.chm158924803
f1_keywords:
- vbawd10.chm158924803
ms.prod: word
api_name:
- Word.CaptionLabel.IncludeChapterNumber
ms.assetid: 6b9c58e6-bb66-1334-278f-aa447103414e
ms.date: 06/08/2017
---


# CaptionLabel.IncludeChapterNumber Property (Word)

 **True** if a chapter number is included with page numbers or a caption label. Read/write **Boolean** .


## Syntax

 _expression_ . **IncludeChapterNumber**

 _expression_ Required. A variable that represents a **[CaptionLabel](captionlabel-object-word.md)** object.


## Example

This example adds the chapter number from the Heading 2 style to figure captions, sets the caption numbering style, and then inserts a new figure caption. The document should already contain a Heading 2 style with numbering.


```vb
With CaptionLabels(wdCaptionFigure) 
 .IncludeChapterNumber = True 
 .ChapterStyleLevel = 2 
 .NumberStyle = wdCaptionNumberStyleUppercaseLetter 
End With 
Selection.InsertCaption Label:="Figure", Title:=": History"
```


## See also


#### Concepts


[CaptionLabel Object](captionlabel-object-word.md)

