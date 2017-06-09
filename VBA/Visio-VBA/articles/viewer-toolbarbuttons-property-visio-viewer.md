---
title: Viewer.ToolbarButtons Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.ToolbarButtons
ms.assetid: 7663e0b1-6022-39c3-0268-fba3b287f868
ms.date: 06/08/2017
---


# Viewer.ToolbarButtons Property (Visio Viewer)

Gets or sets the buttons that are available in the toolbar in Microsoft Visio Viewer. Read/write.


## Syntax

 _expression_. **ToolbarButtons**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

String


## Remarks

Use a comma-delimited list of button-name values. See the following table that maps button names (which you can determine by pausing the mouse pointer over a button) to button-name values.

The default list is "About,Sep,ZoomIn,ZoomOut,ZoomWidth,ZoomPage,Zoom100,Zoom,Sep,OpenInVisio,Sep,Props,Layers,Reviewing,Sep,Help".



|**Toolbar Button**|**Button-Name Value**|
|:-----|:-----|
| **About Microsoft Office Visio Viewer**|About|
| **Zoom In**|ZoomIn|
| **Zoom Out**|ZoomOut|
| **Zoom Width**|ZoomWidth|
| **Zoom Page**|ZoomPage|
| **Zoom 100%**|Zoom100|
| **Zoom**|Zoom|
| **Open in Microsoft Office Visio**|OpenInVisio|
| **Properties and Settings**|Props|
| **Layer Settings**|Layers|
| **Markup Settings**|Reviewing|
| **Help**|Help|
| **Separator**|Sep|
| **First Page**|FirstPage|
| **Previous Page**|PrevPage|
| **Next Page**|NextPage|
| **Last Page**|LastPage|
| **Go To Page**|GoToPage|

## Example

The following code shows how to display the names of the current toolbar buttons in Visio Viewer in the  **Immediate** window.


```vb
Debug.Print vsoViewer.ToolbarButtons
```


