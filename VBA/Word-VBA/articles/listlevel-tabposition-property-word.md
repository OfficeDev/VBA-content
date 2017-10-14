---
title: ListLevel.TabPosition Property (Word)
keywords: vbawd10.chm160235528
f1_keywords:
- vbawd10.chm160235528
ms.prod: word
api_name:
- Word.ListLevel.TabPosition
ms.assetid: 36b73a32-4e8a-f6f5-75d0-55f1ad411055
ms.date: 06/08/2017
---


# ListLevel.TabPosition Property (Word)

Returns or sets the tab position for the specified  **ListLevel** object. Read/write **Single** .


## Syntax

 _expression_ . **TabPosition**

 _expression_ An expression that returns a **[ListLevel](listlevel-object-word.md)** object.


## Remarks

Because the  **ListLevel** object does not have a default tab setting, the **TabPosition** property always returns a value of 999999 or **wdUndefined** , unless you set the property to a value.


## Example

This example sets each list level number so that it is indented 0.5 inch (36 points) from the previous level, and the tab is set at 0.25 inch (18 points) from the number.


```vb
r = 0 
For Each lev In ListGalleries(wdOutlineNumberGallery) _ 
 .ListTemplates(1).ListLevels 
 lev.Alignment = wdListLevelAlignLeft 
 lev.NumberPosition = r 
 lev.TrailingCharacter = wdTrailingTab 
 lev.TabPosition = r + 18 
 r = r + 36 
Next lev
```

This example sets the variable myltemp to the first numbered list template, and then it sets the tab position at one inch. The list template is then applied to the selection.




```vb
Set myltemp = ListGalleries(wdNumberGallery).ListTemplates(1) 
myltemp.ListLevels(1).TabPosition = InchesToPoints(1) 
Selection.Range.ListFormat.ApplyListTemplate ListTemplate:=myltemp
```


## See also


#### Concepts


[ListLevel Object](listlevel-object-word.md)

