---
title: Viewer.ReviewerCount Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.ReviewerCount
ms.assetid: 5ab6cae5-ea59-bb72-1fb2-04aebc5ae5cc
ms.date: 06/08/2017
---


# Viewer.ReviewerCount Property (Visio Viewer)

Gets the count of reviewers in the current document open in Microsoft Visio Viewer. Read-only.


## Syntax

 _expression_. **ReviewerCount**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

 **Long**


## Remarks

The collection of reviewers is one-based, so the index of the first reviewer in the collection is 1.


## Example

The following code gets the number of reviewers in the drawing open in Visio Viewer and displays it in the  **Immediate** window.


```vb
Debug.Print vsoViewer.ReviewerCount
```


