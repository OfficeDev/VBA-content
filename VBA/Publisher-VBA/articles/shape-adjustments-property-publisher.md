---
title: Shape.Adjustments Property (Publisher)
keywords: vbapb10.chm2228273
f1_keywords:
- vbapb10.chm2228273
ms.prod: publisher
api_name:
- Publisher.Shape.Adjustments
ms.assetid: 14794cba-c671-51e3-0aac-52e885a4ba7f
ms.date: 06/08/2017
---


# Shape.Adjustments Property (Publisher)

Returns an  **[Adjustments](adjustments-object-publisher.md)** collection representing all adjustment handles for the specified **Shape** or **ShapeRange** object.


## Syntax

 _expression_. **Adjustments**

 _expression_A variable that represents a  **Shape** object.


## Remarks

Adjustment handles correspond to Microsoft Publisher shape sliders.


## Example

This example takes the number of adjustments for a given shape range and assigns it to a variable.


```vb
Public Sub Counter() 
 
 Dim intCount as Integer 
 
 ' A Shape must be in the active publication and selected. 
 intCount = Publisher.ActiveDocument.Selection _ 
 .ShapeRange(1).Adjustments.Count 
 
End Sub
```


