---
title: Shape.LayerCount Property (Visio)
keywords: vis_sdr.chm11213815
f1_keywords:
- vis_sdr.chm11213815
ms.prod: visio
api_name:
- Visio.Shape.LayerCount
ms.assetid: 0ebcdf53-ebf3-8e26-236f-086f2c9f3c08
ms.date: 06/08/2017
---


# Shape.LayerCount Property (Visio)

Returns the number of layers to which a shape is assigned. Read-only.


## Syntax

 _expression_ . **LayerCount**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Integer


## Remarks

A shape is assigned to zero or more layers.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **LayerCount** property to determine the number of layers to which a shape is assigned. It also uses the **Layers** property to add a layer to a drawing and the **Name** property to get the name of the current layer.


```vb
 
Public Sub LayerCount_Example() 
 
 Dim vsoPage As Visio.Page 
 Dim vsoShape As Visio.Shape 
 Dim vsoLayers As Visio.Layers 
 Dim vsoLayer As Visio.Layer 
 
 If ActiveDocument Is Nothing Then 
 Documents.Add ("") 
 End If 
 
 Set vsoPage = ActivePage 
 If vsoPage Is Nothing Then 
 Set vsoPage = ActiveDocument.Pages(1) 
 End If 
 
 'Draw a rectangle. 
 Set vsoShape = vsoPage.DrawRectangle(1, 5, 5, 1) 
 
 'Get the Layers collection. 
 Set vsoLayers = vsoPage.Layers 
 
 'Create a layer named ExampleLayer1 and add the shape to it. 
 Set vsoLayer = vsoLayers.Add("ExampleLayer1") 
 vsoLayer.Add vsoShape, 1 
 
 'Create a layer named ExampleLayer2 and add the shape to it. 
 Set vsoLayer = vsoLayers.Add("ExampleLayer2") 
 vsoLayer.Add vsoShape, 1 
 
 'Verify that the shape has been assigned to 2 layers. 
 Debug.Print "The shape is assigned to " &; vsoShape.LayerCount &; " layers." 
 
 'Get a reference to the first layer. 
 Set vsoLayer = vsoShape.Layer(1) 
 
 'Verify by using the Name property. 
 Debug.Print "Current layer name is """ &; vsoLayer.Name &; ".""" 
 
End Sub
```


