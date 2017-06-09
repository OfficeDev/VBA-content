---
title: FillFormat.TextureName Property (PowerPoint)
keywords: vbapp10.chm552020
f1_keywords:
- vbapp10.chm552020
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.TextureName
ms.assetid: c8ca47e7-90c8-50b8-2e7e-29e56ec0f70e
ms.date: 06/08/2017
---


# FillFormat.TextureName Property (PowerPoint)

Returns the name of the custom texture file for the specified fill. Read-only.


## Syntax

 _expression_. **TextureName**

 _expression_ A variable that represents a **FillFormat** object.


### Return Value

String


## Remarks

This property is read-only. Use the  **[UserTextured](fillformat-usertextured-method-powerpoint.md)** method to set the texture file for the fill.


## Example

This example adds an oval to myDocument. If shape one on myDocument has a user-defined textured fill, the new oval will have the same fill as shape one. If shape one has any other type of fill, the new oval will have a green marble fill.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes
    Set newFill = .AddShape(msoShapeOval, 0, 0, 200, 90).Fill
    With .Item(1).Fill
        If .Type = msoFillTextured And _
                .TextureType = msoTextureUserDefined Then
            newFill.UserTextured .TextureName
        Else
            newFill.PresetTextured msoTextureGreenMarble
        End If
    End With
End With
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-powerpoint.md)

