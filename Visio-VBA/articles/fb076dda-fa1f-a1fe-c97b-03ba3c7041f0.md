
# Shape.Layer Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns the layer to which a shape is assigned. Read-only.

## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Layer**( **_Index_**)

 _expression_A variable that represents a  **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Integer**|The ordinal of the layer to get.|

### Return Value

Layer


## Remarks
<a name="sectionSection1"> </a>

If a shape is assigned to three layers, the valid indexes that can be passed to its  **Layer** property are 1 through 3.

To get the number of layers to which a shape is assigned, use the  **LayerCount** property.


## Example
<a name="sectionSection2"> </a>

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Layer** property to get a reference to a particular layer. It also uses the **LayerCount** property to determine the number of layers to which a shape is assigned and the **Name** property to get the name of the current layer.


```
 
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
 Debug.Print "The rectangle is assigned to " &amp; vsoShape.LayerCount &amp; " layers." 
 
 'Get a reference to the first layer. 
 Set vsoLayer = vsoShape.Layer(1) 
 
 'Verify by using the Name property. 
 Debug.Print "Current vsoLayer name is """ &amp; vsoLayer.Name &amp; ".""" 
 
End Sub 

```

