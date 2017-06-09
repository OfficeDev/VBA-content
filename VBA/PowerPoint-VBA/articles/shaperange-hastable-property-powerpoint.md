---
title: ShapeRange.HasTable Property (PowerPoint)
keywords: vbapp10.chm548068
f1_keywords:
- vbapp10.chm548068
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.HasTable
ms.assetid: aaf47e4f-0315-2311-e9c5-68a12d36235c
ms.date: 06/08/2017
---


# ShapeRange.HasTable Property (PowerPoint)

Returns whether the specified shape is a table. Read-only.


## Syntax

 _expression_. **HasTable**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

MsoTriState


## Remarks

The value of the  **HasTable** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified shape is not a table.|
|**msoTrue**| The specified shape is a table.|

## Example

This example checks the currently selected shape to see if it is a table. If it is, the code sets the width of column one to one inch (72 points).


```vb
With ActiveWindow.Selection.ShapeRange

    If .HasTable = msoTrue Then

       .Table.Columns(1).Width = 72

    End If

End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

