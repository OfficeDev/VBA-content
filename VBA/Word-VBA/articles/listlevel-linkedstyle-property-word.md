---
title: ListLevel.LinkedStyle Property (Word)
keywords: vbawd10.chm160235531
f1_keywords:
- vbawd10.chm160235531
ms.prod: word
api_name:
- Word.ListLevel.LinkedStyle
ms.assetid: 11a48d9a-87fa-6cc6-8614-deb35775d6b5
ms.date: 06/08/2017
---


# ListLevel.LinkedStyle Property (Word)

Returns or sets the name of the style that's linked to the specified  **ListLevel** object. Read/write **String** .


## Syntax

 _expression_ . **LinkedStyle**

 _expression_ An expression that returns a **[ListLevel](listlevel-object-word.md)** object.


## Example

This example sets the variable myListTemp to the first list template (excluding None) on the  **Outline Numbered** tab in the **Bullets and Numbering** dialog box ( **Format** menu). Each level in the list has a matching heading style linked to it.


```vb
Set myListTemp = _ 
 ListGalleries(wdOutlineNumberGallery).ListTemplates(1) 
For Each mylevel In myListTemp.ListLevels 
 mylevel.LinkedStyle = "Heading " &; mylevel.index 
Next mylevel
```


## See also


#### Concepts


[ListLevel Object](listlevel-object-word.md)

