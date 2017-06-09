---
title: PageSetup.SectionDirection Property (Word)
keywords: vbawd10.chm158400642
f1_keywords:
- vbawd10.chm158400642
ms.prod: word
api_name:
- Word.PageSetup.SectionDirection
ms.assetid: c1b2eda5-95e5-1a16-139f-c8815c484c86
ms.date: 06/08/2017
---


# PageSetup.SectionDirection Property (Word)

Returns or sets the reading order and alignment for the specified sections. Read/write  **WdSectionDirection** .


## Syntax

 _expression_ . **SectionDirection**

 _expression_ Required. A variable that represents a **[PageSetup](pagesetup-object-word.md)** object.


## Example

This example sets the direction of the first section in the active document to right-to-left.


```vb
ActiveDocument.Sections(1).PageSetup.SectionDirection = _ 
 wdSectionDirectionRtl
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-word.md)

