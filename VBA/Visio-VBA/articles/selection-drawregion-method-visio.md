---
title: Selection.DrawRegion Method (Visio)
keywords: vis_sdr.chm11116225
f1_keywords:
- vis_sdr.chm11116225
ms.prod: visio
api_name:
- Visio.Selection.DrawRegion
ms.assetid: 3c3a04d9-a275-a73e-8325-eadd3cae1999
ms.date: 06/08/2017
---


# Selection.DrawRegion Method (Visio)

Draws a new shape that represents the region containing a given point.


## Syntax

 _expression_ . **DrawRegion**( **_Tolerance_** , **_Flags_** , **_x_** , **_y_** , **_ResultsMaster_** )

 _expression_ A variable that represents a **Selection** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Tolerance_|Required| **Double**|Error tolerance when determining the coincidence of points. A distance expressed in internal units in the coordinate space of the  **Selection** object's containing shape; the maximum gap between paths that is tolerated when constructing the boundaries of a region.|
| _Flags_|Required| **Integer**|A constant or integer that specifies how to draw the region.|
| _x_|Optional| **Variant**|x-coordinate in internal units in the coordinate space of the  **Selection** object.|
| _y_|Optional| **Variant**|y-coordinate in internal units in the coordinate space of the  **Selection** object.|
| _ResultsMaster_|Optional| **Variant**|The  **Master** object which the new **Shape** object should be an instance of.|

### Return Value

Shape


## Remarks

The  **DrawRegion** method creates a new **Shape** object from pieces of the paths in the **Selection** object.




- If both  _x_ and _y_ are specified, the resulting shape is the smallest region that contains the point ( _x_, _y_).
    
- In the absence of either  _x_ or _y_, or if the point ( _x_, _y_) is not contained in any region enclosed by the paths of the selected shapes, the result is the union of all the shapes that would have been created by using the  **Fragment** operation.
    
- If no closed region is defined by the selected shapes, the  **DrawRegion** method returns **Nothing** and raises no exception.
    


The  _Flags_ argument can be one or a combination of the following constants declared by the Visio type library in **VisDrawRegionFlags** .



|**Name **|**Value **|**Description **|
|:-----|:-----|:-----|
| **visDrawRegionDeleteInput**|&;H4 |Delete items in selection. |
| **visDrawRegionIgnoreVisible**|&;H20 |Exclude visible geometry. |
| **visDrawRegionIncludeDataGraphics**|&;H40|Include data graphic callout shapes and their sub-shapes. |
| **visDrawRegionIncludeHidden**|&;H10 |Include hidden geometry. |
If the  **DrawRegion** method is passed a _ResultsMaster_ of type VT_EMPTY or VT_ERROR (which is how VBA passes an unspecified optional argument), the new shape is not an instance of a master and the fill, line, and text styles of the new region are set to the document's default styles.

If the  **DrawRegion** method is passed a reference to a **Master** object in _ResultsMaster_ (type VT_UNKNOWN or VT_DISPATCH), the **DrawRegion** method instances that **Master** object and adds geometry computed given the **Selection** object.

The new  **Shape** object has no text other than text already in _ResultsMaster_.


