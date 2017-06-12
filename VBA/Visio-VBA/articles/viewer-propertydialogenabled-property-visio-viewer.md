---
title: Viewer.PropertyDialogEnabled Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.PropertyDialogEnabled
ms.assetid: 66055cb8-535d-16e5-386d-1e7a44faa669
ms.date: 06/08/2017
---


# Viewer.PropertyDialogEnabled Property (Visio Viewer)

Gets or sets a value that indicates whether the  **Properties and Settings** dialog box is available in the user interface for Microsoft Visio Viewer. Read/write.


## Syntax

 _expression_. **PropertyDialogEnabled**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

 **Boolean**


## Remarks

The default is for the  **Properties and Settings** dialog box to be available ( **True**).

When the  **PropertyDialogEnabled** property is set to **False**, clicking  **Properties and Settings** in the toolbar or on the shortcut (right-click) menu has no effect.


## Example

The following code gets a value that indicates whether the  **Properties and Settings** dialog box is available in Visio Viewer.


```vb
Debug.Print vsoViewer.PropertyDialogEnabled
```


