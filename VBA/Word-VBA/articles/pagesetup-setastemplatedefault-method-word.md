---
title: PageSetup.SetAsTemplateDefault Method (Word)
keywords: vbawd10.chm158400714
f1_keywords:
- vbawd10.chm158400714
ms.prod: word
api_name:
- Word.PageSetup.SetAsTemplateDefault
ms.assetid: 3938fd43-6850-d991-be89-b59ef744ac97
ms.date: 06/08/2017
---


# PageSetup.SetAsTemplateDefault Method (Word)

Sets the specified page setup formatting as the default for the active document and all new documents based on the active template.


## Syntax

 _expression_ . **SetAsTemplateDefault**

 _expression_ Required. A variable that represents a **[PageSetup](pagesetup-object-word.md)** object.


## Example

This example changes the left and right margin settings for the active document and then sets the page setup formatting as the default.


```vb
With ActiveDocument.PageSetup 
 .LeftMargin = InchesToPoints(1) 
 .RightMargin = InchesToPoints(1) 
 .SetAsTemplateDefault 
End With
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-word.md)

