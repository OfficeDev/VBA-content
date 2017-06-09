---
title: ListLevel.StartAt Property (Word)
keywords: vbawd10.chm160235530
f1_keywords:
- vbawd10.chm160235530
ms.prod: word
api_name:
- Word.ListLevel.StartAt
ms.assetid: 7331be7c-952e-cd3e-82c0-06712082e6d7
ms.date: 06/08/2017
---


# ListLevel.StartAt Property (Word)

Returns or sets the starting number for the specified  **ListLevel** object. Read/write **Long** .


## Syntax

 _expression_ . **StartAt**

 _expression_ An expression that returns a **[ListLevel](listlevel-object-word.md)** object.


## Example

This example sets the number style and starting number for the third outline-numbered list template. Because the style uses uppercase letters and the starting number is 4, the first letter is D.


```vb
Set mylev = ListGalleries(wdOutlineNumberGallery) _ 
 .ListTemplates(3).ListLevels(1) 
With mylev 
 .NumberStyle = wdListNumberStyleUppercaseLetter 
 .StartAt = 4 
End With
```


## See also


#### Concepts


[ListLevel Object](listlevel-object-word.md)

