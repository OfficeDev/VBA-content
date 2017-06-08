---
title: ListLevel.TextPosition Property (Word)
keywords: vbawd10.chm160235527
f1_keywords:
- vbawd10.chm160235527
ms.prod: word
api_name:
- Word.ListLevel.TextPosition
ms.assetid: ed0ea5ae-d017-a0a8-be0a-cd49015e3bfb
ms.date: 06/08/2017
---


# ListLevel.TextPosition Property (Word)

Returns or sets the position (in points) for the second line of wrapping text for the specified  **ListLevel** object. Read/write **Single** .


## Syntax

 _expression_ . **TextPosition**

 _expression_ An expression that returns a **[ListLevel](listlevel-object-word.md)** object.


## Example

This example sets the indentation for all levels of the first outline-numbered list template. Each list level number is indented 0.5 inch (36 points) from the previous level, the tab is set at 0.25 inch (18 points) from the number, and wrapping text is indented 0.25 inch (18 points) from the number.


```vb
r = 0 
For Each lev In ListGalleries(wdOutlineNumberGallery) _ 
 .ListTemplates(1).ListLevels 
 lev.Alignment = wdListLevelAlignLeft 
 lev.NumberPosition = r 
 lev.TrailingCharacter = wdTrailingTab 
 lev.TabPosition = r + 18 
 lev.TextPosition = r + 18 
 r = r + 36 
Next lev
```


## See also


#### Concepts


[ListLevel Object](listlevel-object-word.md)

