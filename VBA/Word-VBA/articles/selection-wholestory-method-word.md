---
title: Selection.WholeStory Method (Word)
keywords: vbawd10.chm158663180
f1_keywords:
- vbawd10.chm158663180
ms.prod: word
api_name:
- Word.Selection.WholeStory
ms.assetid: ecd50a78-ecbd-75a9-2565-31d7e6ac449a
ms.date: 06/08/2017
---


# Selection.WholeStory Method (Word)

Expands a selection to include the entire story.


## Syntax

 _expression_ . **WholeStory**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

The following instructions, where  _objSel_ is a valid **Selection** object, are functionally equivalent:


```
objSel.WholeStory 
objSel.Expand Unit:=wdStory
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

