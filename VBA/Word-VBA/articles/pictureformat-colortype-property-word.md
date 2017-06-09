---
title: PictureFormat.ColorType Property (Word)
keywords: vbawd10.chm164298853
f1_keywords:
- vbawd10.chm164298853
ms.prod: word
api_name:
- Word.PictureFormat.ColorType
ms.assetid: f4596bf7-4602-385d-61c0-0aed87aaf420
ms.date: 06/08/2017
---


# PictureFormat.ColorType Property (Word)

Returns or sets the type of color transformation applied to the specified picture or OLE object. Read/write  **MsoPictureColorType** .


## Syntax

 _expression_ . **ColorType**

 _expression_ Required. A variable that represents a **[PictureFormat](pictureformat-object-word.md)** object.


## Example

This example sets the color transformation to grayscale for the first shape on the active document. The first shape must be either a picture or an OLE object.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 
 
docActive.Shapes(1).PictureFormat.ColorType = _ 
 msoPictureGrayScale
```


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-word.md)

