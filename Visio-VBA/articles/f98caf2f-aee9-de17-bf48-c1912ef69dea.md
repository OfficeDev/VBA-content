
# InvisibleApp.DefaultZoomBehavior Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Determines the zoom behavior for all new Microsoft Visio documents and drawing windows. Read/write.


## Syntax

 _expression_. **DefaultZoomBehavior**

 _expression_A variable that represents an  **InvisibleApp** object.


### Return Value

VisZoomBehavior


## Remarks

 To set zoom behavior for an existing document, or for a particular window, use the **ZoomBehavior** property of the document and window, respectively.

The following constants declared by the Visio type library in  **VisZoomBehavior** are valid values for the **DefaultZoomBehavior** property.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visZoomNone**|0|Undefined zoom behavior; use the zoom behavior of the document or application|
| **visZoomInPlaceContainer**|1|The container performs the zoom. This is the default.|
| **visZoomVisio**|2|Visio performs the zoom. |
| **visZoomVisioExact**|4|Visio zooms when open in place; Visio does not adjust the zoom level|



 **Note**  The default behavior ( **visZoomInPlaceContainer**) is different from the behavior used in Microsoft Visio 2002, but is the same as that in earlier versions of Visio. To replicate the behavior seen in Microsoft Visio 2002, set this value to  **visZoomVisio**.

If this value is set to the default,  **visZoomInPlaceContainer**, Visio uses the container's  **IOleCommandTarget** interface to perform the zoom and forces a fit-to-window zoom within the in-place window. If the container does not support **IOleCommandTarget**, no zooming occurs.

