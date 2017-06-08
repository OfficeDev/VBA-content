---
title: Shape.GetWidth Method (Publisher)
keywords: vbapb10.chm2228249
f1_keywords:
- vbapb10.chm2228249
ms.prod: publisher
api_name:
- Publisher.Shape.GetWidth
ms.assetid: 9df33329-c37b-82f5-93b4-fc4752ee907e
ms.date: 06/08/2017
---


# Shape.GetWidth Method (Publisher)

Returns the width of the shape or shape range as a  **Single** in the specified units.


## Syntax

 _expression_. **GetWidth**( **_Unit_**)

 _expression_A variable that represents a  **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Unit|Required| **PbUnitType**|The units in which to return the width.|

### Return Value

Single


## Remarks

The Unit parameter can be one of the  **[PbUnitType](pbunittype-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.

Use the  **[GetHeight](shape-getheight-method-publisher.md)** method to return the height of a shape or shape range.


## Example

The following example displays the height and width in inches (to the nearest hundredth) of the shape range consisting of all the shapes on the first page of the active publication.


```vb
With ActiveDocument.Pages(1).Shapes.Range 
 MsgBox "Height of all shapes: " _ 
 &; Format(.GetHeight(Unit:=pbUnitInch), "0.00") _ 
 &; " in" &; vbCr _ 
 &; "Width of all shapes: " _ 
 &; Format(.GetWidth(Unit:=pbUnitInch), "0.00") _ 
 &; " in" 
End With
```


