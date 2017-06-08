---
title: Viewer.CustomPropertyCount Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.CustomPropertyCount
ms.assetid: d036b187-5cb7-87da-b136-fdaa6624b2d4
ms.date: 06/08/2017
---


# Viewer.CustomPropertyCount Property (Visio Viewer)

Gets the count of shape data items (custom properties) assigned to the specified shape in Microsoft Visio Viewer. Read-only.


## Syntax

 _expression_. **CustomPropertyCount**( **_ShapeIndex_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ShapeIndex|Required| **Long**|The index of the shape that contains the shape data (custom properties).|

### Return Value

 **Long**


## Remarks

In versions of Visio prior to Microsoft Office Visio 2007, shape data items were called custom properties.


## Example

The following code gets the count of shape data items assigned to the first shape in the collection of shapes on the current page in Visio Viewer.


```vb
Dim intShapeData As Integer

intShapeData = vsoViewer.CustomPropertyCount(1)

Debug.Print intShapeData


```


