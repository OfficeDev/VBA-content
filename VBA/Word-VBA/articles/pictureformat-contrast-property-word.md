---
title: PictureFormat.Contrast Property (Word)
keywords: vbawd10.chm164298854
f1_keywords:
- vbawd10.chm164298854
ms.prod: word
api_name:
- Word.PictureFormat.Contrast
ms.assetid: 43b91fc2-9a6d-c4d2-c68a-1c8f8a1a00b7
ms.date: 06/08/2017
---


# PictureFormat.Contrast Property (Word)

Returns or sets the contrast for the specified picture or OLE object. The value for this property must be a number from 0.0 (the least contrast) to 1.0 (the greatest contrast). Read/write  **Single** .


## Syntax

 _expression_ . **Contrast**

 _expression_ A variable that represents a **[PictureFormat](pictureformat-object-word.md)** object.


## Example

This example sets the contrast for the first shape on the active document. The first shape must be either a picture or an OLE object.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 
 
docActive.Shapes(1).PictureFormat.Contrast = 0.8
```


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-word.md)

