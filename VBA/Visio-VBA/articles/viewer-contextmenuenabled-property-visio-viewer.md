---
title: Viewer.ContextMenuEnabled Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.ContextMenuEnabled
ms.assetid: 0617d59d-59ed-4012-7dc5-d0aa8be8d110
ms.date: 06/08/2017
---


# Viewer.ContextMenuEnabled Property (Visio Viewer)

Gets or sets a value that indicates whether the shortcut (right-click) menu is enabled in Microsoft Visio Viewer. Read/write.


## Syntax

 _expression_. **ContextMenuEnabled**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

 **Boolean**


## Remarks

The default is for the shortcut menu to be enabled ( **True**).

If the  **ContextMenuEnabled** property is set to **False**, there is no way for the user to access that menu. However, all of the commands on the menu are available on the toolbar.


## Example

The following code disables the shortcut menu in Visio Viewer.


```vb
vsoViewer.ContextMenuEnabled = False
```


