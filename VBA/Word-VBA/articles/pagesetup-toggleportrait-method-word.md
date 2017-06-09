---
title: PageSetup.TogglePortrait Method (Word)
keywords: vbawd10.chm158400713
f1_keywords:
- vbawd10.chm158400713
ms.prod: word
api_name:
- Word.PageSetup.TogglePortrait
ms.assetid: 184fe44c-deb5-3183-742e-88f0c990e62a
ms.date: 06/08/2017
---


# PageSetup.TogglePortrait Method (Word)

Switches between portrait and landscape page orientations for a document or section.


## Syntax

 _expression_ . **TogglePortrait**

 _expression_ Required. A variable that represents a **[PageSetup](pagesetup-object-word.md)** object.


## Remarks

If the specified sections have different page orientations, an error occurs.


## Example

This example changes the page orientation for the active document.


```vb
ActiveDocument.PageSetup.TogglePortrait
```

This example changes the page orientation for all the sections in the selection. If the initial orientation of each section is not the same as the orientation of the other sections, an error occurs.




```
Selection.PageSetup.TogglePortrait
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-word.md)

