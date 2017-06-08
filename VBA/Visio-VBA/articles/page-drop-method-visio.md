---
title: Page.Drop Method (Visio)
keywords: vis_sdr.chm10916235
f1_keywords:
- vis_sdr.chm10916235
ms.prod: visio
api_name:
- Visio.Page.Drop
ms.assetid: 015615a8-fe64-5b76-39ba-ef7ed62e6846
ms.date: 06/08/2017
---


# Page.Drop Method (Visio)

Creates one or more new  **Shape** objects by dropping an object onto a receiving object such as a master, drawing page, shape, or group.


## Syntax

 _expression_ . **Drop**( **_ObjectToDrop_** , **_xPos_** , **_yPos_** )

 _expression_ A variable that represents a **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectToDrop_|Required| **[UNKNOWN]**|The object or selection to drop. While this is typically a Visio object such as a  **Master** , **Shape** , or **Selection** object, it can be any OLE object that provides an **IDataObject** interface.|
| _xPos_|Required| **Double**|The x-coordinate at which to place the center of the shape's width or PinX.|
| _yPos_|Required| **Double**|The y-coordinate at which to place the center of the shape's height or PinY.|

### Return Value

Shape


## Remarks

Using the  **Drop** method is similar to moving a shape with the mouse. The object dropped (ObjectToDrop) can be a master or a shape on the drawing page.

To add a shape to a group or on a drawing page, apply the  **Drop** method to a **Shape** or **Page** object, respectively. The center of the shape's width-height box is positioned at the specified coordinates, and a **Shape** object that represents the shape that is created is returned. When applying this method to a **Shape** object, make sure that the **Shape** object represents a group.

If ObjectToDrop is a  **Master** , the pin of the master is dropped at the specified coordinates. A master's pin is often, but not necessarily, at its center of rotation.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this method maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVPage.Drop(object, double, double)**
    

## Example

The following example shows how to use the  **Drop** method to drop shapes onto **Page** and **Shape** objects.


```vb
 
Public Sub Drop_Example() 
  
    Dim vsoShape1 As Visio.Shape  
    Dim vsoShape2 As Visio.Shape  
    Dim vsoShape3 As Visio.Shape  
    Dim vsoGroupShape As Visio.Shape  
    Dim vsoSubShape As Visio.Shape  
    Dim vsoSelection As Visio.Selection 
  
    Set vsoShape1 = ActivePage.DrawRectangle(1, 2, 2, 1)  
    Set vsoShape2 = ActivePage.DrawRectangle(1, 4, 2, 3)  
 
    'Drop a shape on the page.  
    Set vsoShape3 = ActivePage.Drop(vsoShape1, 3.5, 3.5)  
 
    'Make sure only one shape is selected to start.  
    Set vsoSelection = ActiveWindow.Selection 
    vsoSelection.Select vsoShape1, visDeselectAll + visSelect  
    vsoSelection.Select vsoShape2, visSelect  
 
    'Create a group shape.  
    Set vsoGroupShape = vsoSelection.Group  
 
    'Drop a shape on the group shape to create a new subshape.  
    Set vsoSubShape = vsoGroupShape.Drop(vsoShape3, 1, 2)  
 
End Sub
```


