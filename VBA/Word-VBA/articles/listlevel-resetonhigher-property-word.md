---
title: ListLevel.ResetOnHigher Property (Word)
keywords: vbawd10.chm160235533
f1_keywords:
- vbawd10.chm160235533
ms.prod: word
api_name:
- Word.ListLevel.ResetOnHigher
ms.assetid: 6623910d-94ac-62c7-8af5-5cc32ef9c88f
ms.date: 06/08/2017
---


# ListLevel.ResetOnHigher Property (Word)

Sets or returns the list level that must appear before the specified list level restarts numbering at 1. Read/write  **Long** .


## Syntax

 _expression_ . **ResetOnHigher**

 _expression_ An expression that returns a **[ListLevel](listlevel-object-word.md)** object.


## Remarks

The  **ResetOnHigher** property returns **False** if the numbering continues sequentially each time the list level appears.

This feature allows lists to be interleaved, maintaining numeric sequence. You cannot set the  **ResetOnHigher** property of a list level to a value greater than or equal to its index in the **[ListLevels](listlevels-object-word.md)** collection.


## Example

This example sets each of the nine list levels in the first outline-numbered list template to continue its sequential numbering whenever that level is used.


```vb
For Each li In _ 
 ListGalleries(wdOutlineNumberGallery) _ 
 .ListTemplates(1).ListLevels 
 li.ResetOnHigher = False 
Next li
```


## See also


#### Concepts


[ListLevel Object](listlevel-object-word.md)

