---
title: ListLevel.TrailingCharacter Property (Word)
keywords: vbawd10.chm160235523
f1_keywords:
- vbawd10.chm160235523
ms.prod: word
api_name:
- Word.ListLevel.TrailingCharacter
ms.assetid: 9f64d28c-4409-6278-e20e-baaea1d03ce7
ms.date: 06/08/2017
---


# ListLevel.TrailingCharacter Property (Word)

Returns or sets the character inserted after the number for the specified list level. Read/write  **WdTrailingCharacter** .


## Syntax

 _expression_ . **TrailingCharacter**

 _expression_ Required. A variable that represents a **[ListLevel](listlevel-object-word.md)** object.


## Example

This example sets the number and text alignment for each level of the sixth outline-numbered list template. The number for each level is followed by a space.


```vb
r = 0 
For Each lev In ListGalleries(wdOutlineNumberGallery) _ 
 .ListTemplates(6).ListLevels 
 lev.Alignment = wdListLevelAlignLeft 
 lev.NumberPosition = r 
 lev.TextPosition = r 
 lev.TrailingCharacter = wdTrailingSpace 
 r = r + 18 
Next lev
```


## See also


#### Concepts


[ListLevel Object](listlevel-object-word.md)

