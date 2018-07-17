---
title: FillFormat.UserTextured Method (Word)
keywords: vbawd10.chm164102162
f1_keywords:
- vbawd10.chm164102162
ms.prod: word
api_name:
- Word.FillFormat.UserTextured
ms.assetid: 4407df38-2660-5974-eadb-e30fe292ef11
ms.date: 06/08/2017
---


# FillFormat.UserTextured Method (Word)

Fills the specified shape with small tiles of an image.


## Syntax

 _expression_ . **UserTextured**( **_TextureFile_** )

 _expression_ Required. A variable that represents a **[FillFormat](fillformat-object-word.md)** object.


## Remarks

If you want to fill the shape with one large image, use the  **[UserPicture](fillformat-userpicture-method-word.md)** method.


## Example

This example adds two rectangles to the active document. The rectangle on the left is filled with one large image of the picture in Tiles.bmp; the rectangle on the right is filled with many small tiles of the picture in Tiles.bmp


```vb
Sub Texture() 
 With ActiveDocument.Shapes 
 .AddShape(msoShapeRectangle, 0, 0, 200, 100).Fill _ 
 .UserPicture "C:\Windows\Tiles.bmp" 
 .AddShape(msoShapeRectangle, 300, 0, 200, 100).Fill _ 
 .UserTextured "C:\Windows\Tiles.bmp" 
 End With 
End Sub
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-word.md)

