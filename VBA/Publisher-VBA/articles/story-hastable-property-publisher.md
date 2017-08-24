---
title: Story.HasTable Property (Publisher)
keywords: vbapb10.chm5832707
f1_keywords:
- vbapb10.chm5832707
ms.prod: publisher
api_name:
- Publisher.Story.HasTable
ms.assetid: bc4912e2-f521-c6b5-b5a6-a49952014966
ms.date: 06/08/2017
---


# Story.HasTable Property (Publisher)

Returns  **msoTrue** if the shape represents a **TableFrame** object or **msoFalse** if the shape represents any other object type. Read-only.


## Syntax

 _expression_. **HasTable**

 _expression_A variable that represents a  **Story** object.


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


