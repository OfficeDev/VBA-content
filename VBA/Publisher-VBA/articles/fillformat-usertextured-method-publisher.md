---
title: FillFormat.UserTextured Method (Publisher)
keywords: vbapb10.chm2359320
f1_keywords:
- vbapb10.chm2359320
ms.prod: publisher
api_name:
- Publisher.FillFormat.UserTextured
ms.assetid: fe1a1e06-8bdc-8022-6d4b-6f320f587baf
ms.date: 06/08/2017
---


# FillFormat.UserTextured Method (Publisher)

Fills the specified shape with small tiles of an image.


## Syntax

 _expression_. **UserTextured**( **_TextureFile_**)

 _expression_A variable that represents a  **FillFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|TextureFile|Required| **String**|The name of the texture file.|

## Remarks

To fill the shape with one large image, use the  **[UserPicture](fillformat-userpicture-method-publisher.md)** method.


## Example

This example adds two rectangles to the active publication. The rectangle on the left is filled with one large image of a picture; the rectangle on the right is filled with many small tiles of the same picture. (Note that PathToFile must be replaced with a valid file path for this example to work.)


```vb
With ActiveDocument.Pages(1).Shapes 
 .AddShape(Type:=msoShapeRectangle, _ 
 Left:=0, Top:=0, Width:=200, Height:=100).Fill _ 
 .UserPicture PictureFile:="PathToFile" 
 .AddShape(Type:=msoShapeRectangle, _ 
 Left:=300, Top:=0, Width:=200, Height:=100).Fill _ 
 .UserTextured TextureFile:="PathToFile" 
End With 

```


