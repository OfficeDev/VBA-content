---
title: Crop.ShapeHeight Property (Office)
ms.prod: office
api_name:
- Office.Crop.ShapeHeight
ms.assetid: ccc352b8-2c1e-3116-6f50-a709760189fe
ms.date: 06/08/2017
---


# Crop.ShapeHeight Property (Office)

Gets or sets the height of a shape that is used to crop an image. Read/write


## Syntax

 _expression_. **ShapeHeight**

 _expression_ An expression that returns a **Crop** object.


### Return Value

Single


## Example

The following example inserts a 200 x 200 image into a PowerPoint presentation approximately in the center of the slide. It then resizes the image inside the frame to 100 x 100. The image frame stays at 200 x 200. The code then adds a square (the default shape) just above and to the right of the image, essentially cropping the lower left corner of the image.


```
Sub CropImage() 
 ActivePresentation.Slides(1).Shapes.AddPicture "c:\myImage.png", msoFalse, msoTrue, 250,150, 200, 200 
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.PictureHeight = 100 
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.PictureWidth = 100 
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.PictureOffsetX = 0 
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.PictureOffsetY = 0 
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.ShapeHeight = 100 
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.ShapeWidth = 100 
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.ShapeLeft = 330 
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.ShapeTop = 170 
End Sub 

```


## See also


#### Concepts


[Crop Object](crop-object-office.md)
#### Other resources


[Crop Object Members](crop-members-office.md)

