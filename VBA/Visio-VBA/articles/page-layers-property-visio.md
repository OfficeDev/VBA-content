---
title: Page.Layers Property (Visio)
keywords: vis_sdr.chm10913820
f1_keywords:
- vis_sdr.chm10913820
ms.prod: visio
api_name:
- Visio.Page.Layers
ms.assetid: 62e3aae6-1cb1-695e-81ec-eabdd6b44ef9
ms.date: 06/08/2017
---


# Page.Layers Property (Visio)

Returns the  **Layers** collection of an object. Read-only.


## Syntax

 _expression_ . **Layers**

 _expression_ A variable that represents a **Page** object.


### Return Value

Layers


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Layers** property to add shapes to layers. It also uses the **Layer** property to get a reference to a particular layer, the **LayerCount** property to determine the number of layers to which a shape is assigned, and the **Name** property to get the name of the current layer.


```vb
 
Public Sub Layers_Example() 
 
 Dim vsoPage As Visio.Page 
 Dim vsoShape As Visio.Shape 
 Dim vsoLayer As Visio.Layer 
 Dim vsoLayers As Visio.Layers 
 
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


