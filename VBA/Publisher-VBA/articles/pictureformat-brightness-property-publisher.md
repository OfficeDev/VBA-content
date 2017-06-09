---
title: PictureFormat.Brightness Property (Publisher)
keywords: vbapb10.chm3604736
f1_keywords:
- vbapb10.chm3604736
ms.prod: publisher
api_name:
- Publisher.PictureFormat.Brightness
ms.assetid: bed1cd25-faee-6fb9-4bb3-5bdaf148b62e
ms.date: 06/08/2017
---


# PictureFormat.Brightness Property (Publisher)

Returns or sets a  **Single** indicating the brightness of the specified picture or OLE object. The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest). Read/write.


## Syntax

 _expression_. **Brightness**

 _expression_A variable that represents a  **PictureFormat** object.


### Return Value

Single


## Remarks

Use the  **[IncrementBrightness](pictureformat-incrementbrightness-method-publisher.md)** method to incrementally adjust the brightness from its current level.


## Example

This example sets the brightness for the first shape in the active publication. The shape must be either a picture or an OLE object.


```vb
ActiveDocument.Pages(1).Shapes(1).PictureFormat _ 
 .Brightness = 0.3
```


