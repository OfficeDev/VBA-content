---
title: Layer.Add Method (Visio)
keywords: vis_sdr.chm11816670
f1_keywords:
- vis_sdr.chm11816670
ms.prod: visio
api_name:
- Visio.Layer.Add
ms.assetid: 6e1bd140-426e-cb2d-b883-17ac07117137
ms.date: 06/08/2017
---


# Layer.Add Method (Visio)

Adds a  **Shape** object to a **Layer** object.


## Syntax

 _expression_ . **Add**( **_SheetObject_** , **_fPresMems_** )

 _expression_ A variable that represents a **Layer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SheetObject_|Required| **[IVSHAPE]**|The new  **Shape** object added to the **Layer** object.|
| _fPresMems_|Required| **Integer**|Zero to remove subshapes from any previous layer assignments; non-zero to preserve layer assignments.|

### Return Value

Nothing


## Remarks

If the shape is a group and  _fPresMems_ is non-zero, the component shapes of the group retain their current layer assignments and are also added to this layer. If _fPresMems_ is zero, the component shapes are reassigned to this layer and lose their current layer assignments.


## Example

The following example shows how to use the  **Add** method to add **Shape** objects to a **Layer** object. When the **Shape** object you add to a **Layer** object is a group shape, use the _fPresMems_ argument of the **Add** method to specify whether the component shapes of the group retain or lose their previous layer assignments. If the shape you add is not a group shape, the _fPresMems_ argument has no effect, but is still required.

In the example, two new layers are created. Two rectangle shapes are drawn and then added to the first layer. Subsequently, the rectangles are grouped into a group shape. The group shape is then selected and duplicated, and the duplicate group shapes are added to the second layer in two different ways.

Layer assignments of the component shapes of  _vsoShapeGroup2_ are retained by passing a non-zero value for the _fPresMems_ argument of the **Add** method, but previous layer assignments of the component shapes of _vsoShapeGroup1_ are lost when zero is passed to the **Add** method for that argument. As a result, the component shapes of _vsoShapeGroup1_ are assigned only to _vsoLayer2_ , while the components of _vsoShapeGroup2_ are assigned to both _vsoLayer1_ and _vsoLayer2_ .




```vb
Public Sub AddShapesToLayer_Example() 
 
 Dim vsoDocument As Visio.Document 
 Dim vsoPages As Visio.Pages 
 Dim vsoPage As Visio.Page 
 Dim vsoLayers As Visio.Layers 
 Dim vsoLayer1 As Visio.Layer 
 Dim vsoLayer2 As Visio.Layer 
 Dim vsoShape1 As Visio.Shape 
 Dim vsoShape2 As Visio.Shape 
 Dim vsoShapeGroup1 As Visio.Shape 
 Dim vsoShapeGroup2 As Visio.Shape 
 
 'Add a Document object based on the Basic Diagram template. 
 Set vsoDocument = Documents.Add("Basic Diagram.vst") 
 
 'Get the Pages collection and add a page to the collection. 
 Set vsoPages = vsoDocument.Pages 
 Set vsoPage = vsoPages.Add 
 
 'Get the Layers collection and add two layers 
 'to the collection. 
 Set vsoLayers = vsoPage.Layers 
 Set vsoLayer1 = vsoLayers.Add("MyLayer") 
 Set vsoLayer2 = vsoLayers.Add("MySecondLayer") 
 
 'Draw two rectangles. 
 Set vsoShape1 = vsoPage.DrawRectangle(3, 3, 5, 6) 
 Set vsoShape2 = vsoPage.DrawRectangle(4, 4, 6, 7) 
 
 'Assign each rectangle to the first layer. 
 vsoLayer1.Add vsoShape1, 0 
 vsoLayer1.Add vsoShape2, 0 
 
 'Select the two rectangles and group them. 
 ActiveWindow.SelectAll 
 ActiveWindow.Selection.Group 
 
 'Duplicate the group and set each group as a Shape object. 
 Set vsoShapeGroup1 = vsoPage.Shapes(1) 
 vsoShapeGroup1.Duplicate 
 Set vsoShapeGroup2 = vsoPage.Shapes(2) 
 
 'Add the first grouped shape to the second layer. 
 'This group's component shapes are added to the layer 
 'but lose their previous layer assignment. 
 vsoLayer2.Add vsoShapeGroup1, 0 
 
 'Add the second grouped shape to the second layer. 
 'This group's component shapes are added to the layer 
 'but retain their previous layer assignment. 
 vsoLayer2.Add vsoShapeGroup2, 1 
 
End Sub
```


