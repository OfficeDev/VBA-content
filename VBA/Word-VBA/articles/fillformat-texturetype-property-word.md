---
title: FillFormat.TextureType Property (Word)
keywords: vbawd10.chm164102254
f1_keywords:
- vbawd10.chm164102254
ms.prod: word
api_name:
- Word.FillFormat.TextureType
ms.assetid: 5254a20e-477d-c69e-7296-129deb1e08e0
ms.date: 06/08/2017
---


# FillFormat.TextureType Property (Word)

Returns the texture type for the specified fill. Read-only  **MsoTextureType** .


## Syntax

 _expression_ . **TextureType**

 _expression_ An expression that represents a **[FillFormat](fillformat-object-word.md)** object.


## Remarks

This property is read-only. Use the  **[PresetTextured](fillformat-presettextured-method-word.md)** , **[UserPicture](fillformat-userpicture-method-word.md)** , or **[UserTextured](fillformat-usertextured-method-word.md)** method to set the texture type for the fill.


## Example

This example changes the fill for all shapes in the active document with a custom textured fill to a canvas fill.


```vb
For Each s In ActiveDocument.Shapes 
 With s.Fill 
 If .TextureType = msoTextureUserDefined Then 
 .PresetTextured msoTextureCanvas 
 End If 
 End With 
Next
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-word.md)

