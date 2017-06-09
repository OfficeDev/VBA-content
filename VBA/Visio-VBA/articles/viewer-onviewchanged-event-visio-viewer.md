---
title: Viewer.OnViewChanged Event (Visio Viewer)
ms.prod: visio
api_name:
- Visio.OnViewChanged
ms.assetid: 4d402263-91e1-434c-5f0d-ae7febdc72ab
ms.date: 06/08/2017
---


# Viewer.OnViewChanged Event (Visio Viewer)

Occurs when the view of the current page is changed in Microsoft Visio Viewer.


## Syntax

 _expression_. **OnViewChanged**( **_PageXAtViewCenter_**,  **_PageYAtViewCenter_**,  **_ZoomFactor_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|PageXAtViewCenter|Required| **Double**|The x-coordinate, in drawing units, of the center of the page.|
|PageYAtViewCenter|Required| **Double**|The y-coordinate, in drawing units, of the center of the page.|
|ZoomFactor|Required| **Double**|The factor by which the zoom (the page size) is multiplied. |

### Return Value

Nothing


## Remarks

The page view consists of the center point of the page, expressed in x-y page coordinates, with the origin of the coordinate system at the lower left corner of the page; and the zoom factor, expressed as a numerical percentage, ranging from 1% to 400%.

You can get the current page view in Visio Viewer by using the  **[GetPageView](viewer-getpageview-method-visio-viewer.md)** method, and you can set the page view programmatically by using the **[SetPageView](viewer-setpageview-method-visio-viewer.md)** method.


## Example

The following code shows how to use the  **OnViewChanged** event to display the new page-view data in the **Immediate** window.


```vb
Private Sub vsoViewer_OnViewChanged(ByVal PageXAtViewCenter As Double, ByVal PageYAtViewCenter As Double, ByVal ZoomFactor As Double)

    Dim dblXPoint As Double

    Dim dblYPoint As Double

    Dim dblZoomFactor As Double



    vsoViewer.GetPageView dblXPoint, dblYPoint, dblZoomFactor



    Debug.Print "New x-coordinate is:"; dblXPoint

    Debug.Print "New y-coordinate is:"; dblYPoint

    Debug.Print "New zoom factor is:"; dblZoomFactor

End Sub
```


