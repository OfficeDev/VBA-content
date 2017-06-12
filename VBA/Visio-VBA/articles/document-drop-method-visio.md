---
title: Document.Drop Method (Visio)
keywords: vis_sdr.chm10516235
f1_keywords:
- vis_sdr.chm10516235
ms.prod: visio
api_name:
- Visio.Document.Drop
ms.assetid: 1e6b2d14-71c2-4adc-a9d7-ec123b2b7f31
ms.date: 06/08/2017
---


# Document.Drop Method (Visio)

Creates a new  **Master** object by dropping an object onto a receiving object such as a stencil or document, or the **Masters** or **MasterShortcuts** collection.


## Syntax

 _expression_ . **Drop**( **_ObjectToDrop_** , **_xPos_** , **_yPos_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectToDrop_|Required| **[UNKNOWN]**|The object to drop. While this is typically a Visio object such as a  **Master** , **Shape** , or **Selection** object, it can be any OLE object that provides an **IDataObject** interface.|
| _xPos_|Required| **Integer**|The x-coordinate at which to place the center of the shape's width or PinX.|
| _yPos_|Required| **Integer**|The y-coordinate at which to place the center of the shape's height or PinY.|

### Return Value

Master


## Remarks

Using the  **Drop** method is similar to moving a shape with the mouse. The object dropped ( _ObjectToDrop_) can be a master or a shape on the drawing page.

If  _ObjectToDrop_ is a **Master** , the pin of the master is dropped at the specified coordinates. A master's pin is often, but not necessarily, at its center of rotation.

To create a new master in a stencil, apply the  **Drop** method to a **Document** object that represents a stencil (the stencil must be opened as an original or a copy rather than read-only). In this case, the _xPos_ and _yPos_ arguments are ignored, and the new master that is created is returned.


## Example

The following example shows how to use the  **Drop** method to create a master by dropping a shape onto a **Document** object.


```vb
 
Public Sub Drop_Example() 
 
    Dim vsoShape As Visio.Shape  
    Dim vsoMaster As Visio.Master  
 
    Set vsoShape = ActivePage.DrawRectangle(1, 2, 2, 1)  
 
    'Create a master in the document.  
    'The master appears on the document stencil.  
    Set vsoMaster = ActiveDocument.Drop(vsoShape,  0, 0)  
 
End Sub
```


