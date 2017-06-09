---
title: Work with Panes and Views
keywords: vbapp10.chm5275171
f1_keywords:
- vbapp10.chm5275171
ms.prod: powerpoint
ms.assetid: 13301fb3-b22e-19d9-b181-ee006e05dd90
ms.date: 06/08/2017
---


# Work with Panes and Views

## Changing the Active View

You can return or set the current view in the active document window with the  **ViewType** property. This example changes the view in the active document window to slide view.


```vb
ActiveWindow.ViewType = ppViewSlide
```


## Changing Panes in Normal View

In normal view, you can use the  **ViewType** property with the active **Pane** object to return the active pane. The **ViewType** property returns one of the following **PpViewType** constants, identifying the active pane: **ppViewNotesPage**, **ppViewOutline**, or **ppViewSlide**. All other views have only one pane and the **ViewType** property returns the same **PpViewType** constant value as the active document window.

You can activate a pane by setting the  **ViewType** property or by using the **Activate** method. This example returns the value of the **ViewType** property to identify the active view and active pane. If the active view is normal view and the active pane is the notes pane, then the slide pane is activated with the **Activate** method.




```vb
With ActiveWindow
    If .ViewType = ppViewNormal and _
            .ActivePane.ViewType = ppViewNotesPage Then
        .Panes(2).Activate
    End If
End With
```


## Resizing panes

You can use the  **SplitHorizontal** and the **SplitVertical** properties to reposition the pane dividers in normal view to the specified percentage of the available document window. This resizes the panes on either side of the divider. The maximum value of these properties is always less than 100% because the slide pane has a minimum size that depends on a 10% zoom level. This example sets the percentage of the available document window height that the slide pane occupies to 65 percent, leaving the notes pane at 35 percent.


```vb
ActiveWindow.SplitVertical = 65
```


