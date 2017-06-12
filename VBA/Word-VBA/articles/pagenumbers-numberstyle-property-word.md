---
title: PageNumbers.NumberStyle Property (Word)
keywords: vbawd10.chm159776770
f1_keywords:
- vbawd10.chm159776770
ms.prod: word
api_name:
- Word.PageNumbers.NumberStyle
ms.assetid: 5a7a3101-3b16-a107-8790-3666fa7fba54
ms.date: 06/08/2017
---


# PageNumbers.NumberStyle Property (Word)

Returns or sets a  **[WdPageNumberStyle](wdpagenumberstyle-enumeration-word.md)** constant that represents the number style. Read/write.


## Syntax

 _expression_ . **NumberStyle**

 _expression_ Required. An expression that returns a **[PageNumbers](pagenumbers-object-word.md)** object.


## Example

This example formats the page numbers in the active document's footer as lowercase roman numerals.


```vb
For Each sec In ActiveDocument.Sections 
 sec.Footers(wdHeaderFooterPrimary).PageNumbers _ 
 .NumberStyle = wdPageNumberStyleLowercaseRoman 
Next sec
```


## See also


#### Concepts


[PageNumbers Collection Object](pagenumbers-object-word.md)

