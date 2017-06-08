---
title: FillFormat.TextureType Property (Publisher)
keywords: vbapb10.chm2359568
f1_keywords:
- vbapb10.chm2359568
ms.prod: publisher
api_name:
- Publisher.FillFormat.TextureType
ms.assetid: 08f3b0a1-97a3-bdbf-25b4-93e05938d607
ms.date: 06/08/2017
---


# FillFormat.TextureType Property (Publisher)

Returns an  **MsoTextureType** constant indicating the texture type for the specified fill. Read-only.


## Syntax

 _expression_. **TextureType**

 _expression_A variable that represents a  **FillFormat** object.


### Return Value

MsoTextureType


## Remarks

This property is read-only. Use the  [PresetTextured](fillformat-presettextured-method-publisher.md)or  **[UserTextured](fillformat-usertextured-method-publisher.md)** method to set the texture type for the fill.

The property value can be one of the  **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoTexturePreset**| The fill uses a preset texture type.|
| **msoTextureTypeMixed**|Indicates a combination of texture types for the specified shape range..|
| **msoTextureUserDefined**|The fill uses a user-defined texture type.|

## Example

This example applies a canvas texture to the fill for all shapes on the first page of the active publication that currently have fills with a user-defined texture.


```vb
Dim shpLoop As Shape 
 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 With shpLoop.Fill 
 If .TextureType = msoTextureUserDefined Then 
 .PresetTextured _ 
 PresetTexture:=msoTextureCanvas 
 End If 
 End With 
Next shpLoop
```


