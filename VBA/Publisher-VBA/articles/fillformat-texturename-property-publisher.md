---
title: FillFormat.TextureName Property (Publisher)
keywords: vbapb10.chm2359561
f1_keywords:
- vbapb10.chm2359561
ms.prod: publisher
api_name:
- Publisher.FillFormat.TextureName
ms.assetid: 237a85ff-018d-f6b7-e94b-32e85fce65ab
ms.date: 06/08/2017
---


# FillFormat.TextureName Property (Publisher)

Returns a  **String** indicating the name of the custom texture file for the specified fill. Read-only.


## Syntax

 _expression_. **TextureName**

 _expression_A variable that represents a  **FillFormat** object.


### Return Value

String


## Remarks

Use the  **[UserTextured](fillformat-usertextured-method-publisher.md)** method to set the texture file for the fill.


## Example

This example adds an oval to the active publication. If shape one on the active publication has a fill with a user-defined texture, the new oval will have the same fill as shape one. If shape one has any other type of fill, the new oval will have a green marble fill.


```vb
Dim ffNew As FillFormat 
 
With ActiveDocument.Pages(1).Shapes 
 Set ffNew = .AddShape(Type:=msoShapeOval, _ 
 Left:=0, Top:=0, Width:=200, Height:=90).Fill 
 
 With .Item(1).Fill 
 If .Type = msoFillTextured And _ 
 .TextureType = msoTextureUserDefined Then 
 ffNew.UserTextured _ 
 TextureFile:=.TextureName 
 Else 
 ffNew.PresetTextured _ 
 PresetTexture:=msoTextureGreenMarble 
 End If 
 End With 
End With 

```


