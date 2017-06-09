---
title: ShapeRange.GetWidth Method (Publisher)
keywords: vbapb10.chm2293785
f1_keywords:
- vbapb10.chm2293785
ms.prod: publisher
api_name:
- Publisher.ShapeRange.GetWidth
ms.assetid: a15d1b50-289a-8b02-e090-0f0a9637980a
ms.date: 06/08/2017
---


# ShapeRange.GetWidth Method (Publisher)

Returns the width of the shape or shape range as a  **Single** in the specified units. .


## Syntax

 _expression_. **GetWidth**( **_Unit_**)

 _expression_A variable that represents a  **ShapeRange** object.


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


