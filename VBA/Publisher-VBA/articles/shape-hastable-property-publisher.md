---
title: Shape.HasTable Property (Publisher)
keywords: vbapb10.chm2228321
f1_keywords:
- vbapb10.chm2228321
ms.prod: publisher
api_name:
- Publisher.Shape.HasTable
ms.assetid: 6f544d9c-00a4-3047-fbfb-6f1835bbe2c6
ms.date: 06/08/2017
---


# Shape.HasTable Property (Publisher)

Returns  **msoTrue** if the shape represents a **TableFrame** object or **msoFalse** if the shape represents any other object type. Read-only.


## Syntax

 _expression_. **HasTable**

 _expression_A variable that represents a  **Shape** object.


## Remarks

The  **HasTable** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**| The shapes in the range do not represent a **TableFrame** object.|
| **msoTriStateMixed**|Indicates a combination of  **msoTrue** and **msoFalse** for the specified shape range.|
| **msoTrue**|The shapes in the range represent a  **TableFrame** object.|

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


