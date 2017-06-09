---
title: Shape.Layer Property (Visio)
keywords: vis_sdr.chm11213810
f1_keywords:
- vis_sdr.chm11213810
ms.prod: visio
api_name:
- Visio.Shape.Layer
ms.assetid: fb076dda-fa1f-a1fe-c97b-03ba3c7041f0
ms.date: 06/08/2017
---


# Shape.Layer Property (Visio)

Returns the layer to which a shape is assigned. Read-only.


## Syntax

 _expression_ . **Layer**( **_Index_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Integer**|The ordinal of the layer to get.|

### Return Value

Layer


## Remarks

If a shape is assigned to three layers, the valid indexes that can be passed to its  **Layer** property are 1 through 3.

To get the number of layers to which a shape is assigned, use the  **LayerCount** property.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Layer** property to get a reference to a particular layer. It also uses the **LayerCount** property to determine the number of layers to which a shape is assigned and the **Name** property to get the name of the current layer.


```vb
 
Public Sub Layer_Example() 
 
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
 Debug.Print "The rectangle is assigned to " &; vsoShape.LayerCount &; " layers." 
 
 'Get a reference to the first layer. 
 Set vsoLayer = vsoShape.Layer(1) 
 
 'Verify by using the Name property. 
 Debug.Print "Current vsoLayer name is """ &; vsoLayer.Name &; ".""" 
 
End Sub
```


