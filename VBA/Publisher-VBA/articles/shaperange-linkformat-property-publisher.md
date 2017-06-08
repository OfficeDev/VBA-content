---
title: ShapeRange.LinkFormat Property (Publisher)
keywords: vbapb10.chm2293862
f1_keywords:
- vbapb10.chm2293862
ms.prod: publisher
api_name:
- Publisher.ShapeRange.LinkFormat
ms.assetid: 1f0add8d-7baa-65f0-e82b-a047a7bc0507
ms.date: 06/08/2017
---


# ShapeRange.LinkFormat Property (Publisher)

Returns a  [LinkFormat](linkformat-object-publisher.md)object that contains the properties that are unique to linked OLE objects. Read-only.


## Syntax

 _expression_. **LinkFormat**

 _expression_A variable that represents a  **ShapeRange** object.


## Example

This example updates the links between any OLE objects on page one in the active publication and their source files.


```vb
Dim sh As Shape 
 
For Each sh In ActiveDocument.Pages(1).Shapes 
 If sh.Type = pbLinkedOLEObject Then 
 With sh.LinkFormat 
 .Update 
 End With 
 End If 
Next
```


