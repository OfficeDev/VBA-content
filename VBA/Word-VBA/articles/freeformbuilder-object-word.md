---
title: FreeformBuilder Object (Word)
keywords: vbawd10.chm2505
f1_keywords:
- vbawd10.chm2505
ms.prod: word
api_name:
- Word.FreeformBuilder
ms.assetid: 31e89628-4b50-ff72-ce3d-dc7c161dad3e
ms.date: 06/08/2017
---


# FreeformBuilder Object (Word)

Represents the geometry of a freeform while it is being built.


## Remarks

Use the  **BuildFreeform** method of the **[Shapes](shapes-object-word.md)** or **[CanvasShapes](canvasshapes-object-word.md)** object to return a **FreeformBuilder** object. Use the **[AddNodes](freeformbuilder-addnodes-method-word.md)** method to add nodes to the freeform. Use the **[ConvertToShape](freeformbuilder-converttoshape-method-word.md)** method to create the shape defined in the **FreeformBuilder** object and add it to the **Shapes** collection. The following example adds a freeform with four segments to the active document.


```vb
With ActiveDocument.Shapes _ 
 .BuildFreeform(msoEditingCorner, 360, 200) 
 .AddNodes msoSegmentCurve, msoEditingCorner, _ 
 380, 230, 400, 250, 450, 300 
 .AddNodes msoSegmentCurve, msoEditingAuto, 480, 200 
 .AddNodes msoSegmentLine, msoEditingAuto, 480, 400 
 .AddNodes msoSegmentLine, msoEditingAuto, 360, 200 
 .ConvertToShape 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


