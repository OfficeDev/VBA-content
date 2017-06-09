---
title: TextEffectFormat.NormalizedHeight Property (Publisher)
keywords: vbapb10.chm3735814
f1_keywords:
- vbapb10.chm3735814
ms.prod: publisher
api_name:
- Publisher.TextEffectFormat.NormalizedHeight
ms.assetid: 2b62fe23-9204-7449-1d4e-73e73def5df0
ms.date: 06/08/2017
---


# TextEffectFormat.NormalizedHeight Property (Publisher)

Specifies whether all characters (both uppercase and lowercase) in the specified WordArt are the same height. Read/write.


## Syntax

 _expression_. **NormalizedHeight**

 _expression_A variable that represents a  **TextEffectFormat** object.


### Return Value

MsoTriState


## Remarks

The  **NormalizedHeight** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**| Characters in the specified WordArt object are not all the same height.|
| **msoTrue**| Characters in the specified WordArt object are all the same height.|

## Example

This example creates a new WordArt shape on the first page of the active publication and then sets each character in the shape to be the same height.


```vb
Sub SetNormalHeight() 
 With ActiveDocument.Pages(1).Shapes.AddTextEffect _ 
 (PresetTextEffect:=msoTextEffect10, _ 
 text:="Test WordArt Shape", FontName:="Snap ITC", _ 
 FontSize:=30, FontBold:=msoFalse, FontItalic:=msoFalse, _ 
 Left:=36, Top:=36).TextEffect 
 .NormalizedHeight = msoTrue 
 End With 
End Sub
```


