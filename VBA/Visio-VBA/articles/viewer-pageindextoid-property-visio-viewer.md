---
title: Viewer.PageIndexToID Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.PageIndexToID
ms.assetid: d354e9d4-1272-2fd1-44dd-5664e94bc6ac
ms.date: 06/08/2017
---


# Viewer.PageIndexToID Property (Visio Viewer)

Gets the ID of the page at the specified index in the collection of pages in the drawing that is open in Microsoft Visio Viewer. Read-only.


## Syntax

 _expression_. **PageIndexToID**( **_PageIndex_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|PageIndex|Required| **Long**|The index of the page whose ID you want to get.|

### Return Value

 **Long**


## Remarks

The collection of pages is one-based, so the index of the first page in the collection is 1.


## Example

The following code shows how to get the ID of the page at index position 1 in the collection of pages in the drawing that is open in Visio Viewer.


```vb
Debug.Print vsoViewer.PageIndexToID(1)
```


