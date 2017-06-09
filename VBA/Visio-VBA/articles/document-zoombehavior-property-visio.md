---
title: Document.ZoomBehavior Property (Visio)
keywords: vis_sdr.chm10551465
f1_keywords:
- vis_sdr.chm10551465
ms.prod: visio
api_name:
- Visio.Document.ZoomBehavior
ms.assetid: 5507fc17-957a-ab7f-d15f-43ad3e8327c6
ms.date: 06/08/2017
---


# Document.ZoomBehavior Property (Visio)

Determines the zoom behavior for a Microsoft Visio document or window. Read/write.


## Syntax

 _expression_ . **ZoomBehavior**

 _expression_ A variable that represents a **Document** object.


### Return Value

VisZoomBehavior


## Remarks

To set the zoom behavior for all new documents and windows, use the  **DefaultZoomBehavior** property.

The following constants declared by the Visio type library in  **VisZoomBehavior** are valid values for **ZoomBehavior** .



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visZoomNone**|0|Undefined zoom behavior; use the zoom behavior of the document or application. This is the default.|
| **visZoomInPlaceContainer**|1|The container performs the zoom.|
| **visZoomVisio**|2|Visio performs the zoom. |
| **visZoomVisioExact**|4|Visio zooms when open in place; Visio does not adjust the zoom level|
If  **ZoomBehavior** is set to **visZoomVisio** , Visio adjusts the zoom level to certain discrete values, for example 50% or 100%, to optimize the appearance of the page rulers and grid, and of snap behavior.

If  **ZoomBehavior** is set to **visZoomInPlaceContainer** , Visio uses the container's **IOleCommandTarget** interface to perform the zoom and forces a fit-to-window zoom within the in-place window. If the container does not support **IOleCommandTarget** , no zooming occurs.

If  **ZoomBehavior** is set to **visZoomVisioExact** , you can set the zoom to any value, either by using the **Window.Zoom** property or by using the **Zoom** slider in the Visio user interface.




 **Note**  The default behavior ( **visZoomNone** ) is different from the behavior used in versions earlier than Visio 2002. (In Visio 2002, the default was **visZoomVisio** .) To replicate the behavior seen in earlier versions, set this value to **visZoomInPlaceContainer** .


## Example

The following procedure shows how to set the  **Document.ZoomBehavior** property to the value that replicates Visio 2000 behavior.


```vb
Sub ZoomBehavior_Example() 
 
 ActiveDocument.ZoomBehavior = visZoomInPlaceContainer 
 
End Sub
```


