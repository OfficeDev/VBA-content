---
title: Layer.Name Property (Visio)
keywords: vis_sdr.chm11813930
f1_keywords:
- vis_sdr.chm11813930
ms.prod: visio
api_name:
- Visio.Layer.Name
ms.assetid: 4026742c-26e9-2632-dd8b-21021333fcf6
ms.date: 06/08/2017
---


# Layer.Name Property (Visio)

Specifies the name of an object. Read-only.


## Syntax

 _expression_ . **Name**

 _expression_ A variable that represents a **Layer** object.


### Return Value

String


## Remarks




 **Note**  Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **Name** property to get or set a **Hyperlink** , **Layer** , **Master** , **MasterShortcut** , **Page** , **Shape** , **Style** , or **Row** object's local name. Use the **NameU** property to get or set its universal name.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Name** property to display layer names. It also uses the **Layer** property to get a reference to a particular layer and the **LayerCount** property to determine the number of layers to which a shape is assigned.


```vb
 
Public Sub Name_Example() 
 
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
 Debug.Print "The page has " &; vsoShape.LayerCount &; " layers." 
 
 'Get a reference to the first layer. 
 Set vsoLayer = vsoShape.Layer(1) 
 
 'Verify by using the Name property. 
 Debug.Print "Current vsoLayer name is """ &; vsoLayer.Name &; ".""" 
 
End Sub
```


