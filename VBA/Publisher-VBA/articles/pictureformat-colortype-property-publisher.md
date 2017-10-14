---
title: PictureFormat.ColorType Property (Publisher)
keywords: vbapb10.chm3604737
f1_keywords:
- vbapb10.chm3604737
ms.prod: publisher
api_name:
- Publisher.PictureFormat.ColorType
ms.assetid: 439f9eb9-2593-d719-4ef6-0f14d1c7d0f4
ms.date: 06/08/2017
---


# PictureFormat.ColorType Property (Publisher)

Returns or sets an  **MsoPictureColorType** constant indicating the type of color transformation applied to the specified picture or OLE object. Read/write.


## Syntax

 _expression_. **ColorType**

 _expression_A variable that represents a  **PictureFormat** object.


### Return Value

MsoPictureColorType


## Remarks

The  **ColorType** property value can be one of the ** [MsoPictureColorType](http://msdn.microsoft.com/library/d11f2d08-2ac9-6cf4-34b8-7ffaabb5d4ae%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.


## Example

This example sets the color transformation to grayscale for the first shape in the active publication. The shape must be either a picture or an OLE object.


```vb
ActiveDocument.Pages(1).Shapes(1).PictureFormat _ 
 .ColorType = msoPictureGrayScale
```


