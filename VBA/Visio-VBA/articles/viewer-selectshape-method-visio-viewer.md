---
title: Viewer.SelectShape Method (Visio Viewer)
ms.prod: visio
api_name:
- Visio.SelectShape
ms.assetid: 3b3160e3-f4b4-fec2-ae1c-ed274eb69217
ms.date: 06/08/2017
---


# Viewer.SelectShape Method (Visio Viewer)

Selects the specified shape in the drawing that is open in Microsoft Visio Viewer.


## Syntax

 _expression_. **SelectShape**( **_ShapeIndex_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ShapeIndex|Required| **Long**|The index in the collection of shapes of the shape to be selected.|

### Return Value

Nothing


## Remarks

The collection of shapes is one-based, so the first shape in the collection has index number 1.

Passing 0 to the  **SelectShape** method deselects the currently selected shape.


## Example

The following code selects the shape at index position 2 in the collection of shapes in the drawing that is open in Visio Viewer.


```
vsoViewer.SelectShape (2)
```


