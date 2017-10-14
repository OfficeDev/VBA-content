---
title: Viewer.PageIDToIndex Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.PageIDToIndex
ms.assetid: f10470ae-44b8-8817-1c2e-5022f63e8edf
ms.date: 06/08/2017
---


# Viewer.PageIDToIndex Property (Visio Viewer)

Gets the index of the specified page in the collection of pages in the drawing that is open in Microsoft Visio Viewer. Read-only.


## Syntax

 _expression_. **PageIDToIndex**( **_PageID_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|PageID|Required| **Long**|The ID of the page whose index you want to determine.|

### Return Value

 **Long**


## Remarks

The collection of pages is one-based, so the index of the first page in the collection is 1.

If you pass a value for PageID that does not correspond to an actual page ID, the  **PageIDToIndex** property returns 0.


## Example

The following code shows how to get the index of the page in the drawing that is open in Visio Viewer and that has ID 0.


```vb
Debug.Print vsoViewer.PageIDToIndex(0)
```


