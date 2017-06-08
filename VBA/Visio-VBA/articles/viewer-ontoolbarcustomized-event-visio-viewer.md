---
title: Viewer.OnToolbarCustomized Event (Visio Viewer)
ms.prod: visio
api_name:
- Visio.OnToolbarCustomized
ms.assetid: 02796238-7773-309b-a136-1ded2c09f93f
ms.date: 06/08/2017
---


# Viewer.OnToolbarCustomized Event (Visio Viewer)

Occurs when the user customizes the Microsoft Visio Viewer toolbar by adding or removing buttons.


## Syntax

 _expression_. **OnToolbarCustomized**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

Nothing


## Remarks

You can customize the toolbar in Visio Viewer by adding or removing buttons. To do so in the user interface, right-click in the toolbar area, and then click  **Customize**. 

You can customize the toolbar programmatically by using the  **[ToolbarButtons](viewer-toolbarbuttons-property-visio-viewer.md)** property. For the toolbar to be customizable, the **[ToolbarCustomizable](viewer-toolbarcustomizable-property-visio-viewer.md)** property must be set to its default value, **True**.


## Example

The following code shows how to use the  **OnToolbarCustomized** event to display a message in the **Immediate** window when the user customizes the toolbar in Visio Viewer.


```vb
Private Sub vsoViewer_OnToolbarCustomized()

   Debug.Print "The toolbar has been customized!"

End Sub
```


