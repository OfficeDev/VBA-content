---
title: Viewer.Unload Method (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Unload
ms.assetid: 4b746cbf-2f81-b4ef-3f5e-4df93a543292
ms.date: 06/08/2017
---


# Viewer.Unload Method (Visio Viewer)

Unloads the drawing file that is open in Microsoft Visio Viewer.


## Syntax

 _expression_. **Unload**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

Nothing


## Example

The following code prints the name of the drawing that is loaded in Visio Viewer, unloads the drawing, and then prints a blank statement that shows that the document has been unloaded.


```vb
Debug.Print vsoViewer.SRC

vsoViewer.Unload

Debug.Print vsoViewer.SRC
```


