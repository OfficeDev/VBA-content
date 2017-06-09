---
title: TextEffectFormat.PresetShape Property (Publisher)
keywords: vbapb10.chm3735815
f1_keywords:
- vbapb10.chm3735815
ms.prod: publisher
api_name:
- Publisher.TextEffectFormat.PresetShape
ms.assetid: 4e98e606-d26b-aa81-0e19-5b8535ba6df1
ms.date: 06/08/2017
---


# TextEffectFormat.PresetShape Property (Publisher)

Returns or sets an  **MsoPresetTextEffectShape** constant that represents the shape of the specified WordArt. Read/write.


## Syntax

 _expression_. **PresetShape**

 _expression_A variable that represents a  **TextEffectFormat** object.


### Return Value

MsoPresetTextEffectShape


## Remarks

The  **PresetShape** property value can be one of the ** [MsoPresetTextEffectShape](http://msdn.microsoft.com/library/d8b21a00-54af-b2cf-6a1d-4bbaef2aac78%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.


## Example

This example sets the shape of the first shape on the first page of the active publication to a chevron whose center points down. For this example to work the first shape must be a WordArt shape.


```vb
Sub ChangeTextEffect() 
 With ActiveDocument.Pages(1).Shapes(1) 
 If .Type = msoTextEffect Then 
 .TextEffect.PresetShape = msoTextEffectShapeChevronDown 
 End If 
 End With 
End Sub
```


