---
title: TextEffectFormat.PresetTextEffect Property (Publisher)
keywords: vbapb10.chm3735816
f1_keywords:
- vbapb10.chm3735816
ms.prod: publisher
api_name:
- Publisher.TextEffectFormat.PresetTextEffect
ms.assetid: d7ef0995-4560-fea0-b98d-03c8e0c8e65e
ms.date: 06/08/2017
---


# TextEffectFormat.PresetTextEffect Property (Publisher)

Returns or sets an  **MsoPresetTextEffect** constant that represents the style of the specified WordArt. The values for this property correspond to the formats in the **WordArt Gallery** dialog box, numbered from left to right, top to bottom. Read/write.


## Syntax

 _expression_. **PresetTextEffect**

 _expression_A variable that represents a  **TextEffectFormat** object.


### Return Value

MsoPresetTextEffect


## Remarks

The  **PresetTextEffect** property value can be one of the ** [MsoPresetTextEffect](http://msdn.microsoft.com/library/56a7008d-ce2c-f127-56de-851cb8fef44f%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.


## Example

This example sets the text effect style for the first shape on the first page of the active publication. This example assumes that there is at least one shape on the first page of the active publication.


```vb
Sub ChangeTextEffect() 
 With ActiveDocument.Pages(1).Shapes(1) 
 If .Type = msoTextEffect Then 
 .TextEffect.PresetTextEffect = msoTextEffect1 
 End If 
 End With 
End Sub
```


