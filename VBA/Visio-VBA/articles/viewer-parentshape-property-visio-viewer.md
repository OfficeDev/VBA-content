---
title: Viewer.ParentShape Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.ParentShape
ms.assetid: ee6dc52a-86c7-6a8c-c40e-aaad6a1100a5
ms.date: 06/08/2017
---


# Viewer.ParentShape Property (Visio Viewer)

Gets the index number of the parent shape of the specified shape in the drawing that is open in Microsoft Visio Viewer. Read-only.


## Syntax

 _expression_. **ParentShape**( **_ShapeIndex_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ShapeIndex|Required| **Long**|The index of the shape whose parent you want to find.|

### Return Value

 **Long**


## Remarks

The expression "parent shape" refers to the group shape of which the specified shape is a part.

The collection of shapes is one-based, so the index of the first shape in the collection is 1.


## Example

The following code shows how to get the parent of the first shape on the current page in the drawing that is open in Visio Viewer.


```vb
Debug.Print vsoViewer.ParentShape(1)
```


