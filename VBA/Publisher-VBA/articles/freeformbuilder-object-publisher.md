---
title: FreeformBuilder Object (Publisher)
keywords: vbapb10.chm3342335
f1_keywords:
- vbapb10.chm3342335
ms.prod: publisher
api_name:
- Publisher.FreeformBuilder
ms.assetid: 542df9f7-f636-a98e-01de-11005b5797cc
ms.date: 06/08/2017
---


# FreeformBuilder Object (Publisher)

Represents the geometry of a freeform while it is being built.
 


## Example

Use the  **[BuildFreeform](shapes-buildfreeform-method-publisher.md)** method of the **[Shapes](shapes-object-publisher.md)** collection to return a **FreeformBuilder** object. Use the **[AddNodes](freeformbuilder-addnodes-method-publisher.md)** method to add nodes to the freeform. Use the **[ConvertToShape](freeformbuilder-converttoshape-method-publisher.md)** method to create the shape defined in the **FreeformBuilder** object and add it to the **Shapes** collection. The following example adds a freeform with four segments to the active document.
 

 

```
Sub CreateNewFreeFormShape() 
 With ActiveDocument.Pages(1).Shapes.BuildFreeform( _ 
 EditingType:=msoEditingCorner, X1:=360, Y1:=200) 
 .AddNodes SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingCorner, X1:=380, Y1:=230, _ 
 X2:=400, Y2:=250, X3:=450, Y3:=300 
 .AddNodes SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingAuto, X1:=480, Y1:=200 
 .AddNodes SegmentType:=msoSegmentLine, _ 
 EditingType:=msoEditingAuto, X1:=480, Y1:=400 
 .AddNodes SegmentType:=msoSegmentLine, _ 
 EditingType:=msoEditingAuto, X1:=360, Y1:=200 
 .ConvertToShape 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[AddNodes](freeformbuilder-addnodes-method-publisher.md)|
|[ConvertToShape](freeformbuilder-converttoshape-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](freeformbuilder-application-property-publisher.md)|
|[Parent](freeformbuilder-parent-property-publisher.md)|

