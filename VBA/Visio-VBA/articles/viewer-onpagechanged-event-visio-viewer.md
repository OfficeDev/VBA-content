---
title: Viewer.OnPageChanged Event (Visio Viewer)
ms.prod: visio
api_name:
- Visio.OnPageChanged
ms.assetid: de64b0e2-372c-f1c4-15c6-d6c3a4da8368
ms.date: 06/08/2017
---


# Viewer.OnPageChanged Event (Visio Viewer)

Occurs when the active page is changed in Microsoft Visio Viewer.


## Syntax

 _expression_. **OnPageChanged**( **_PageIndex_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|PageIndex|Required| **Long**|The index of the new page.|

### Return Value

Nothing


## Remarks

The collection of pages in the Viewer is one-based, so the index of the first page in the collection is 1. 

You can change the page programmatically in Visio Viewer by setting the value of the  **[CurrentPageIndex](viewer-currentpageindex-property-visio-viewer.md)** property.


## Example

The following code shows how to use the  **OnPageChanged** event to print a message in the **Immediate** window stating that the page has changed and identifying the new page.


```vb
Private Sub vsoViewer_OnPageChanged(ByVal PageIndex As Long)



    Debug.Print "Page changed to"; vsoViewer.CurrentPageIndex

    

End Sub
```


