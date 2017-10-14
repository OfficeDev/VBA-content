---
title: Viewer.SetPageView Method (Visio Viewer)
ms.prod: visio
api_name:
- Visio.SetPageView
ms.assetid: 669c8d29-9793-08a3-05ee-54aab77881bb
ms.date: 06/08/2017
---


# Viewer.SetPageView Method (Visio Viewer)

Sets the position and zoom factor (size) of the drawing page in Microsoft Visio Viewer.


## Syntax

 _expression_. **SetPageView**( **_PageXAtViewCenter_**,  **_PageYAtViewCenter_**,  **_ZoomFactor_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|PageXAtViewCenter|Required| **Double**|The x-coordinate, in drawing-page units, of the center of the page, measured from the lower left corner of the page.|
|PageYAtViewCenter|Required| **Double**|The y-coordinate, in drawing-page units, of the center of the page, measured from the lower left corner of the page.|
|ZoomFactor|Required| **Double**|The factor by which to multiply the zoom (page size).|

### Return Value

Nothing


## Remarks

The page view consists of the center point of the page, expressed in x-y page coordinates, with the origin of the coordinate system at the lower left corner of the page; and the zoom factor, expressed as a numerical percentage, with a range from 1% through 400%.

You can use the  **[GetPageView](viewer-getpageview-method-visio-viewer.md)** method to get the current page-view values.

The  **SetPageView** method sets the coordinates of the point in the page coordinate system that is at the center of the Visio Viewer window. For example, passing 0 for both the x-coordinate and y-coordinate places the lower left corner of the page (the origin of the page's coordinate system) in the center of the Visio Viewer window. If the page is 8 page-units wide by 10 page-units high, passing 4 for PageXAtViewCenter and 5 for PageYAtViewCenter places the center of the page at the center of the Visio Viewer window.

The ZoomFactor parameter value is the factor by which to multiply both dimensions of the page. For example, passing .50 for ZoomFactor makes the page both half as high and half as wide as it was previously.


## Example

The following code sets the center of the page at the center of the Visio Viewer window and halves both the height and width of the page.


```
vsoViewer.SetPageView 4, 5, 0.50
```


