---
title: FillFormat.TextureName Property (Word)
keywords: vbawd10.chm164102253
f1_keywords:
- vbawd10.chm164102253
ms.prod: word
api_name:
- Word.FillFormat.TextureName
ms.assetid: 9eb01e1b-3cd1-16ad-4a7b-a430e27782d9
ms.date: 06/08/2017
---


# FillFormat.TextureName Property (Word)

Returns the name of the custom texture file for the specified fill. Read-only  **String** .


## Syntax

 _expression_ . **TextureName**

 _expression_ An expression that returns a **[FillFormat](fillformat-object-word.md)** object.


## Remarks

Use the  **[UserTextured](fillformat-usertextured-method-word.md)** method to set the texture file for the fill.


## Example

This example adds an oval to the active document. If the second shape in the document has a user-defined textured fill, the new oval will have the same fill as shape two. If shape two has any other type of fill, the new oval will have a green marble fill. This example assumes that the active document already has at least two shapes.


```vb
With ActiveDocument.Shapes 
 Set newFill = .AddShape(msoShapeOval, 0, 0, 200, 90).Fill 
 With .Item(2).Fill 
 If.TextureType = msoTextureUserDefined Then 
 newFill.UserTextured .TextureName 
 Else 
 newFill.PresetTextured msoTextureGreenMarble 
 End If 
 End With 
End With
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-word.md)

