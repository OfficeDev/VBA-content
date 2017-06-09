---
title: Viewer.ReviewerInitial Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.ReviewerInitial
ms.assetid: be7bfc55-d4c0-4d7b-c50d-e6106441ca37
ms.date: 06/08/2017
---


# Viewer.ReviewerInitial Property (Visio Viewer)

Gets the initials of the specified reviewer in the drawing that is open in Microsoft Visio Viewer. Read-only.


## Syntax

 _expression_. **ReviewerInitial**( **_ReviewerIndex_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ReviewerIndex|Required| **Long**|The index of the reviewer in the collection of reviewers.|

### Return Value

String


## Remarks

The collection of reviewers is one-based, so the index of the first reviewer in the collection is 1.


## Example

The following code gets the initials of the reviewer at index position 1 in the collection of reviewers in the drawing that is open in Visio Viewer.


```vb
Debug.Print vsoViewer.ReviewerInitial(1)
```


