---
title: PictureFormat.Brightness Property (Word)
keywords: vbawd10.chm164298852
f1_keywords:
- vbawd10.chm164298852
ms.prod: word
api_name:
- Word.PictureFormat.Brightness
ms.assetid: 385fbf20-db89-e159-31ec-2c9cf3bb5a3a
ms.date: 06/08/2017
---


# PictureFormat.Brightness Property (Word)

Returns or sets the brightness of the specified picture or OLE object. The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest). Read/write  **Single** .


## Syntax

 _expression_ . **Brightness**

 _expression_ A variable that represents a **[PictureFormat](pictureformat-object-word.md)** object.


## Example

This example sets the brightness for the first shape on the active document. The first shape must be either a picture or an OLE object.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 

```


```
docActive.Shapes(1).PictureFormat.Brightness = 0.3
```


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-word.md)

