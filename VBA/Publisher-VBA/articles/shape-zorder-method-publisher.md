---
title: Shape.ZOrder Method (Publisher)
keywords: vbapb10.chm2228272
f1_keywords:
- vbapb10.chm2228272
ms.prod: publisher
api_name:
- Publisher.Shape.ZOrder
ms.assetid: 05143a2b-924e-b5a3-390d-9493627bfa9f
ms.date: 06/08/2017
---


# Shape.ZOrder Method (Publisher)

Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).


## Syntax

 _expression_. **ZOrder**( **_ZOrderCmd_**)

 _expression_A variable that represents a  **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ZOrderCmd|Required| **MsoZOrderCmd**|Specifies where to move the specified shape relative to the other shapes.|

### Return Value

Nothing


## Remarks

The ZOrderCmd parameter can be one of the  **MsoZOrderCmd** constants declared in the Microsoft Office type library and shown in the following table.



| **msoBringForward**|
| **msoBringInFrontOfText**|
| **msoBringToFront**|
| **msoSendBackward**|
| **msoSendBehindText**|
| **msoSendToBack**|
Use the  [ZOrderPosition](shape-zorderposition-property-publisher.md)property to determine a shape's current position in the z-order.


## Example

This example adds an oval to the active publication and then places the oval second from the back in the z-order if there is at least one other shape on the page.


```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeOval, _ 
 Left:=100, Top:=100, Width:=100, Height:=300) 
 While .ZOrderPosition > 2 
 .ZOrder ZOrderCmd:=msoSendBackward 
 Wend 
End With 

```


