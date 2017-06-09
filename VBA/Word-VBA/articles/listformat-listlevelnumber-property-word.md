---
title: ListFormat.ListLevelNumber Property (Word)
keywords: vbawd10.chm163577924
f1_keywords:
- vbawd10.chm163577924
ms.prod: word
api_name:
- Word.ListFormat.ListLevelNumber
ms.assetid: 004c1823-56dd-7a7c-2b0c-8654f0313465
ms.date: 06/08/2017
---


# ListFormat.ListLevelNumber Property (Word)

Returns or sets the list level for the first paragraph in the specified  **ListFormat** object. Read/write **Long** .


## Syntax

 _expression_ . **ListLevelNumber**

 _expression_ Required. A variable that represents a **[ListFormat](listformat-object-word.md)** object.


## Example

This example returns the list level for the third paragraph in the active document.


```
lev = ActiveDocument.Paragraphs(3).Range.ListFormat.ListLevelNumber
```


## See also


#### Concepts


[ListFormat Object](listformat-object-word.md)

