---
title: Viewer.PageName Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.PageName
ms.assetid: 7a23a8da-7763-91fc-777d-fca61e268fe8
ms.date: 06/08/2017
---


# Viewer.PageName Property (Visio Viewer)

Gets the name of the specified page in the drawing that is open in Microsoft Visio Viewer. Read-only.


## Syntax

 _expression_. **PageName**( **_PageIndex_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|PageIndex|Required| **Long**|The index of the page whose name you want to get.|

### Return Value

String


## Remarks

The collection of pages is one-based, so the index of the first page in the collection is 1.

If the local name of the specified page is different from the universal name, the  **PageName** property returns the local name.


## Example

The following code gets the name of the page at index position 1 in the collection of pages in the drawing open in Visio Viewer.


```vb
Debug.Print vsoViewer.PageName(1)
```


