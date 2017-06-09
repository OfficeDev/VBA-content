---
title: ShapeRange.HasTable Property (Publisher)
keywords: vbapb10.chm2293857
f1_keywords:
- vbapb10.chm2293857
ms.prod: publisher
api_name:
- Publisher.ShapeRange.HasTable
ms.assetid: 71ce4980-f5b5-c94c-c29d-32b97cf771fd
ms.date: 06/08/2017
---


# ShapeRange.HasTable Property (Publisher)

Returns  **msoTrue** if the shape represents a **TableFrame** object or **msoFalse** if the shape represents any other object type. Read-only.


## Syntax

 _expression_. **HasTable**

 _expression_A variable that represents a  **ShapeRange** object.


## Example

This example checks the currently selected shape to see if it is a table. If it is, the code sets the width of column one to one inch (72 points).


```vb
Sub IsTable() 
 
 With Application.Selection.ShapeRange 
 If .HasTable = msoTrue Then 
 .Table.Columns(1).Width = 72 
 End If 
 End With 
 
End Sub
```


