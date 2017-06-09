---
title: PictureFormat.Contrast Property (Publisher)
keywords: vbapb10.chm3604738
f1_keywords:
- vbapb10.chm3604738
ms.prod: publisher
api_name:
- Publisher.PictureFormat.Contrast
ms.assetid: f081b7c8-50cc-772b-f3b0-27c215cfebac
ms.date: 06/08/2017
---


# PictureFormat.Contrast Property (Publisher)

Returns or sets a  **Single** indicating the contrast for the specified picture or OLE object. The value for this property must be a number from 0.0 (the least contrast) to 1.0 (the greatest contrast). Read/write.


## Syntax

 _expression_. **Contrast**

 _expression_A variable that represents a  **PictureFormat** object.


### Return Value

Single


## Remarks

Use the  **[IncrementContrast](pictureformat-incrementcontrast-method-publisher.md)** method to incrementally adjust the contrast from its current level.


## Example

This example sets the contrast for the first shape in the active publication. The shape must be either a picture or an OLE object.


```vb
ActiveDocument.Pages(1).Shapes(1).PictureFormat _ 
 .Contrast = 0.8
```


