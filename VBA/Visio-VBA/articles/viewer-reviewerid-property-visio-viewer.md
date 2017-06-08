---
title: Viewer.ReviewerID Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.ReviewerID
ms.assetid: dc6c8175-9cfb-5f31-8572-d7ead88d96a4
ms.date: 06/08/2017
---


# Viewer.ReviewerID Property (Visio Viewer)

Gets the ID of the specified reviewer in the drawing open in Microsoft Visio Viewer. Read-only.


## Syntax

 _expression_. **ReviewerID**( **_ReviewerIndex_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ReviewerIndex|Required| **Long**|The index of the reviewer in the collection of reviewers.|

### Return Value

 **Long**


## Remarks

The collection of reviewers is one-based, so the index of the first reviewer in the collection is 1.


## Example

The following code gets the ID of the reviewer at index position 1 in the drawing open in Visio Viewer.


```vb
Debug.Print vsoViewer.ReviewerID(1)
```


