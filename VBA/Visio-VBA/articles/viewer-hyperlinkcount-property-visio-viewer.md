---
title: Viewer.HyperlinkCount Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.HyperlinkCount
ms.assetid: 06c06812-25a6-779d-3af4-821538493c4f
ms.date: 06/08/2017
---


# Viewer.HyperlinkCount Property (Visio Viewer)

Gets the count of hyperlinks associated with the shape at the specified index in the document open in Microsoft Visio Viewer. Read-only.


## Syntax

 _expression_. **HyperlinkCount**( **_ShapeIndex_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ShapeIndex|Required| **Long**|The index of the specified shape in the collection of shapes in the drawing open in Visio Viewer.|

### Return Value

 **Long**


## Remarks

The collection of shapes is one-based, so the index of the first shape in the collection is 1.


## Example

The following code gets the count of hyperlinks associated with the first shape in the collection of shapes in the drawing open in Visio Viewer.


```vb
Debug.Print vsoViewer.HyperlinkCount(1)
```


