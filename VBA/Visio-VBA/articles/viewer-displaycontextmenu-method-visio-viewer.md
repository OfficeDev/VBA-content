---
title: Viewer.DisplayContextMenu Method (Visio Viewer)
ms.prod: visio
api_name:
- Visio.DisplayContextMenu
ms.assetid: 0aa19901-7bb8-6abe-cbff-4217381af336
ms.date: 06/08/2017
---


# Viewer.DisplayContextMenu Method (Visio Viewer)

Displays the shortcut menu for Microsoft Visio Viewer at the specified screen coordinates, in pixels.


## Syntax

 _expression_. **DisplayContextMenu**( **_ScreenX_**,  **_ScreenY_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ScreenX|Required| **Long**|The x-coordinate, in pixels, of the point where the menu should appear, relative to the origin of the frame of the screen.|
|ScreenY|Required| **Long**|The y-coordinate, in pixels, of the point where the menu should appear, relative to the origin of the frame of the screen.|

### Return Value

Nothing


## Remarks

Use the screenX and screenY parameters to specify the coordinates of the point where you want the shortcut menu to appear, relative to the origin of the frame of the screen. The origin of the screen frame is in the upper left corner.


## Example

The following code specifies that the shortcut menu appear at screen coordinates (300, 300).


```
vsoViewer.DisplayContextMenu(300,300)
```


