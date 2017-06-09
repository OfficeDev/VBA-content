---
title: Viewer.GetPageView Method (Visio Viewer)
ms.prod: visio
api_name:
- Visio.GetPageView
ms.assetid: ad53c016-3a6b-617d-6cfd-93c489f03c69
ms.date: 06/08/2017
---


# Viewer.GetPageView Method (Visio Viewer)

Gets the position and zoom factor (size) of the drawing page in Microsoft Visio Viewer.


## Syntax

 _expression_. **GetPageView**( **_PageXAtViewCenter_**,  **_PageYAtViewCenter_**,  **_ZoomFactor_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|PageXAtViewCenter|Required| **Double**|The x-coordinate of the center of the page, in drawing-page units, measured from the lower left corner of the page.|
|PageYAtViewCenter|Required| **Double**|The y-coordinate of the center of the page, in drawing-page units, measured from the lower left corner of the page.|
|ZoomFactor|Required| **Double**|The factor by which the zoom (page size) is multiplied.|

### Return Value

Nothing


## Remarks

The page view consists of the center point of the page, expressed in x-y page coordinates, with the origin of the coordinate system at the lower left corner of the page; and the zoom factor, expressed as a numerical percentage, with range from 1% to 400%.

You can use the  **[SetPageView](viewer-setpageview-method-visio-viewer.md)** method to set the current page-view values.

The  **GetPageView** method itself returns nothing, but its parameters are all out-parameters. If you pass a variable of type **Double** for each parameter, Visio Viewer returns the respective values of each parameter, as shown in the example in this topic.

The  **GetPageView** method gets the coordinates of the point in the page coordinate system that is at the center of the Visio Viewer window. For example, if Visio Viewer returns 0 for both the x-coordinate and y-coordinate, the lower left corner of the page (the origin of the page's coordinate system) is in the center of the Visio Viewer window. If the page is 8 page-units wide by 10 page-units high, and the center of the page is at the center of the Visio Viewer window, PageXAtViewCenter returns 4 and PageYAtViewCenter returns 5.

The ZoomFactor parameter value is the factor by which both dimensions of the page are multiplied. For example, a ZoomFactor value of .5 means that the page is both half as high and half as wide as it is in the source Visio drawing.


## Example

The following code shows how to get the current position and zoom factor of the page that is open in Visio Viewer.


```vb
 Dim dblXPoint As Double

    Dim dblYPoint As Double

    Dim dblZoomFactor As Double

    vsoViewer.GetPageView dblXPoint, dblYPoint, dblZoomFactor

    Debug.Print "x-coordinate is:"; dblXPoint

    Debug.Print "y-coordinate is:"; dblYPoint

    Debug.Print "Zoom factor is:"; dblZoomFactor
```


