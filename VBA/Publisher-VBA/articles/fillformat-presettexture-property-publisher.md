---
title: FillFormat.PresetTexture Property (Publisher)
keywords: vbapb10.chm2359560
f1_keywords:
- vbapb10.chm2359560
ms.prod: publisher
api_name:
- Publisher.FillFormat.PresetTexture
ms.assetid: c03a9bf3-7378-e82a-9a40-650c5c96fd2a
ms.date: 06/08/2017
---


# FillFormat.PresetTexture Property (Publisher)

Returns an  **MsoPresetTexture** constant that represents the preset texture for the specified fill. Read-only.


## Syntax

 _expression_. **PresetTexture**

 _expression_A variable that represents a  **FillFormat** object.


### Return Value

MsoPresetTexture


## Remarks

The  **PresetTexture** property value can be one of the ** [MsoPresetTexture](http://msdn.microsoft.com/library/fbbc897d-f5db-eb0d-20d9-f6b7e9bbcf4f%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.

Use the  **[PresetTextured](fillformat-presettextured-method-publisher.md)** method to specify the preset texture for the fill.


## Example

This example adds a rectangle to the first page in the active publication and sets its preset texture to match that of the first shape on the page. For the example to work, the first shape must have a preset textured fill.


```vb
Sub SetTexture() 
 Dim texture As MsoPresetTexture 
 With ActiveDocument.Pages(1).Shapes 
 texture = .Item(1).Fill.PresetTexture 
 With .AddShape(Type:=msoShapeRectangle, Left:=250, Top:=72, _ 
 Width:=40, Height:=80) 
 .Fill.PresetTextured PresetTexture:=texture 
 End With 
 End With 
End Sub
```


