---
title: ListLevel.NumberPosition Property (Word)
keywords: vbawd10.chm160235525
f1_keywords:
- vbawd10.chm160235525
ms.prod: word
api_name:
- Word.ListLevel.NumberPosition
ms.assetid: 444df40d-4165-54b9-3456-ca4dfbdb8053
ms.date: 06/08/2017
---


# ListLevel.NumberPosition Property (Word)

Returns or sets the position (in points) of the number or bullet for the specified  **ListLevel** object. Read/write **Single** .


## Syntax

 _expression_ . **NumberPosition**

 _expression_ An expression that returns a **[ListLevel](listlevel-object-word.md)** object.


## Remarks

For each list level, you can set the position of the number or bullet, the position of the tab, and the position of the text that wraps.


## Example

This example sets the indentation for all the levels of the third outline-numbered list template. Each list level is indented 0.25 inch (18 points) more than the preceding level.


```vb
r = 0 
For Each lev In ListGalleries(wdOutlineNumberGallery) _ 
 .ListTemplates(3).ListLevels 
 lev.Alignment = wdListLevelAlignLeft 
 lev.NumberPosition = r 
 r = r + 18 
Next lev
```

This example sets the indent for the first level of the last numbered list template to 0.5 inch.




```vb
With ListGalleries(wdNumberGallery).ListTemplates(7).ListLevels(1) 
 .Alignment = wdListLevelAlignLeft 
 .NumberPosition = InchesToPoints(0.5) 
End With
```


## See also


#### Concepts


[ListLevel Object](listlevel-object-word.md)

