---
title: CanvasShapes.BuildFreeform Method (Word)
keywords: vbawd10.chm7536660
f1_keywords:
- vbawd10.chm7536660
ms.prod: word
api_name:
- Word.CanvasShapes.BuildFreeform
ms.assetid: eb960023-aeda-d272-c9c8-5474fb5867b2
ms.date: 06/08/2017
---


# CanvasShapes.BuildFreeform Method (Word)

Builds a freeform object. Returns a  **[FreeformBuilder](freeformbuilder-object-word.md)** object that represents the freeform as it is being built. .


## Syntax

 _expression_ . **BuildFreeform**( **_EditingType_** , **_X1_** , **_Y1_** )

 _expression_ Required. A variable that represents a **[CanvasShapes](canvasshapes-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EditingType_|Required| **MsoEditingType**|The EditingType parameter can be either  **msoEditingAuto** or **msoEditingCorner** ; cannot be **msoEditingSmooth** or **msoEditingSymmetric** .|
| _X1_|Required| **Single**|The position (in points) of the first node in the freeform drawing relative to the left edge of the document.|
| _Y1_|Required| **Single**|The position (in points) of the first node in the freeform drawing relative to the top of the document.|

## Remarks

Use the  **[AddNodes](freeformbuilder-addnodes-method-word.md)** method to add segments to the freeform. After you have added at least one segment to the freeform, you can use the **ConvertToShape** method to convert the **[FreeformBuilder](freeformbuilder-object-word.md)** object into a **[Shape](shape-object-word.md)** object that has the geometric description you've defined in the **[FreeformBuilder](freeformbuilder-object-word.md)** object.


## Example

This example adds a freeform with five vertices to the active document.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 
With docActive.Shapes.BuildFreeform(msoEditingCorner, 360, 200) 
 .AddNodes msoSegmentCurve, msoEditingCorner, _ 
 380, 230, 400, 250, 450, 300 
 .AddNodes msoSegmentCurve, msoEditingAuto, 480, 200 
 .AddNodes msoSegmentLine, msoEditingAuto, 480, 400 
 .AddNodes msoSegmentLine, msoEditingAuto, 360, 200 
 .ConvertToShape 
End With
```


## See also


#### Concepts


[CanvasShapes Collection](canvasshapes-object-word.md)

