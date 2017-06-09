---
title: FillFormat.TextureType Property (PowerPoint)
keywords: vbapp10.chm552021
f1_keywords:
- vbapp10.chm552021
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.TextureType
ms.assetid: 318e5b2f-7baa-296b-c7ea-0feddb70414c
ms.date: 06/08/2017
---


# FillFormat.TextureType Property (PowerPoint)

Returns the texture type for the specified fill. Read-only.


## Syntax

 _expression_. **TextureType**

 _expression_ A variable that represents a **FillFormat** object.


### Return Value

MsoTextureType


## Remarks

Use the  **[PresetTextured](fillformat-presettextured-method-powerpoint.md)** or **[UserTextured](fillformat-usertextured-method-powerpoint.md)** method to set the texture type for the fill.

The value of the  **TextureType** property can be one of these **MsoTextureType** constants.


||
|:-----|
|**msoTexturePreset**|
|**msoTextureTypeMixed**|
|**msoTextureUserDefined**|

## Example

This example changes the fill to canvas for all shapes on myDocument that have a custom textured fill.


```vb
Set myDocument = ActivePresentation.Slides(1)

For Each s In myDocument.Shapes

    With s.Fill

        If .TextureType = msoTextureUserDefined Then

            .PresetTextured msoTextureCanvas

        End If

    End With

Next

	
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-powerpoint.md)

