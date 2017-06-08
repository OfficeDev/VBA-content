---
title: Viewer.OnMarkupOverlaysVisibleChanged Event (Visio Viewer)
ms.prod: visio
api_name:
- Visio.OnMarkupOverlaysVisibleChanged
ms.assetid: 343f1bd6-07e1-06a0-c707-7b5ca6baa99c
ms.date: 06/08/2017
---


# Viewer.OnMarkupOverlaysVisibleChanged Event (Visio Viewer)

Occurs when the visibility of markup overlays changes in Microsoft Visio Viewer.


## Syntax

 _expression_. **OnMarkupOverlaysVisibleChanged**( **_MarkupOverlaysVisible_** )

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|MarkupOverlaysVisible|Required| **Boolean**|Indicates whether markup overlays are visible in the Visio Viewer user interface.|

## Remarks

You can change the visibility of markup overlays in Visio Viewer by using the  **[MarkupOverlaysVisible](viewer-markupoverlaysvisible-property-visio-viewer.md)** property.

The  **OnMarkupOverlaysVisibleChanged** event occurs when markup overlays for all reviewers in a drawing are set to be visible or not visible. The **[OnReviewerChanged](viewer-onreviewerchanged-event-visio-viewer.md)** event occurs when markup overlays of a specific reviewer are set to be visible or not visible.


## Example

The following code shows how to use the  **OnMarkupOverlaysVisibleChanged** event to display the visibility status of markup overlays in the **Immediate** window.


```vb
Private Sub vsoViewer_OnMarkupOverlaysVisibleChanged(ByVal MarkupOverlaysVisible As Boolean)



    If MarkupOverlaysVisible Then

        Debug.Print "Markup overlays are now visible."

    Else

        Debug.Print "Markup overlays are now not visible."

    End If

        

End Sub
```


